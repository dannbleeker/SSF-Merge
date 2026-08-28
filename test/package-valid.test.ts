import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { DOMParser } from "@xmldom/xmldom";
import { Pkg, resolveTarget } from "../src/core/pptx/pkg.js";
import { prepareBlock } from "../src/core/merge/prepare.js";
import { buildPlan } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import { parseDelimited, toRecordSet } from "../src/core/data/recordset.js";
import { makeDeck, type SlideSpec } from "./fixtures/deck.js";

/**
 * What the engine HANDS OVER, checked as a package rather than as slides.
 *
 * Everything else here tests a decision: does this paragraph merge, does that
 * plan skip the right row. None of it asks the question PowerPoint asks, which
 * is whether the file it is given is a legal OOXML package — and the answer to
 * that one is binary and expensive. A deck that opens as "repaired" has lost
 * whatever PowerPoint decided to drop, silently, in somebody's presentation.
 *
 * The rules below are the package format's, not this project's, so they hold
 * however the engine changes. They are also the shape of the bug this repo has
 * already shipped once: #38 sent the user's whole deck back because the
 * remove-the-rest path was wrong, and a structural check over the output is
 * what would have caught it without a PowerPoint.
 */

const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const parse = (text: string): Document => new DOMParser().parseFromString(text, "text/xml") as unknown as Document;
const els = (d: Document, ns: string, n: string): Element[] => Array.from(d.getElementsByTagNameNS(ns, n));

/** Every way the package could be malformed, as a list of sentences. */
async function problemsIn(bytes: Uint8Array): Promise<string[]> {
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files).filter((n) => !zip.files[n]?.dir);
  const has = (p: string): boolean => names.includes(p);
  const read = async (p: string): Promise<Document> => parse(await zip.file(p)!.async("string"));
  const problems: string[] = [];

  // A relationship pointing at nothing is how a deck opens as "repaired".
  for (const name of names.filter((n) => n.includes("/_rels/") && n.endsWith(".rels"))) {
    const owner = name.replace("/_rels/", "/").replace(/\.rels$/, "");
    const doc = await read(name);
    const rels = els(doc, PKG_REL, "Relationship");
    for (const rel of rels) {
      // An external target is a URL and is not a package path.
      if ((rel.getAttribute("TargetMode") ?? "") === "External") continue;
      const target = rel.getAttribute("Target") ?? "";
      const path = resolveTarget(owner, target);
      if (!has(path))
        problems.push(`${name}: ${rel.getAttribute("Id")} points at ${target}, which is not in the package`);
    }
    const ids = rels.map((r) => r.getAttribute("Id"));
    const dupes = [...new Set(ids.filter((v, i) => ids.indexOf(v) !== i))];
    if (dupes.length) problems.push(`${name}: duplicate rIds ${dupes.join(", ")}`);
  }

  // A part with no content type opens as damaged, and PowerPoint does not say
  // which part it could not classify.
  //
  // The generic rule alone is TOOTHLESS for the parts that matter, and this was
  // measured rather than assumed: a real .pptx declares
  // `<Default Extension="xml" ContentType="application/xml"/>`, so every XML
  // part passes it. Deleting the clone's `addContentTypeOverride` — a defect
  // that would ship every merged slide untyped — left this check green. A slide
  // needs its OWN override naming the slide content type, so those are checked
  // by name.
  const ct = await read("[Content_Types].xml");
  const declared = new Map(
    els(ct, CT_NS, "Override").map((o) => [o.getAttribute("PartName") ?? "", o.getAttribute("ContentType") ?? ""]),
  );
  const defaults = new Set(els(ct, CT_NS, "Default").map((d) => (d.getAttribute("Extension") ?? "").toLowerCase()));
  for (const name of names) {
    if (name === "[Content_Types].xml") continue;
    const ext = (name.split(".").pop() ?? "").toLowerCase();
    if (!declared.has(`/${name}`) && !defaults.has(ext)) problems.push(`no content type declared for ${name}`);
  }
  for (const [part] of declared)
    if (part && !has(part.replace(/^\//, ""))) problems.push(`content type declared for ${part}, which is gone`);

  const NEEDS: [RegExp, string][] = [
    [/^ppt\/slides\/slide\d+\.xml$/, "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"],
    [
      /^ppt\/notesSlides\/notesSlide\d+\.xml$/,
      "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
    ],
  ];
  for (const name of names) {
    for (const [pattern, type] of NEEDS) {
      if (!pattern.test(name)) continue;
      const got = declared.get(`/${name}`);
      if (got === undefined) problems.push(`${name} has no content-type override of its own`);
      else if (got !== type) problems.push(`${name} is declared as ${got}, not ${type}`);
    }
  }

  // The deck's own order, which is what decides the slides a reader sees.
  const pres = await read("ppt/presentation.xml");
  const presRels = await read("ppt/_rels/presentation.xml.rels");
  const target = new Map(
    els(presRels, PKG_REL, "Relationship").map((r) => [r.getAttribute("Id"), r.getAttribute("Target")]),
  );
  const seen = new Set<string>();
  const listed = new Set<string>();
  for (const sldId of els(pres, P_NS, "sldId")) {
    const id = sldId.getAttribute("id") ?? "";
    if (seen.has(id)) problems.push(`two slides share the id ${id}`);
    seen.add(id);
    // The format's range. Outside it PowerPoint rejects the file outright.
    const n = Number(id);
    if (!(n >= 256 && n <= 2147483647)) problems.push(`slide id ${id} is outside the format's range`);
    const rId = sldId.getAttributeNS(R_NS, "id") ?? sldId.getAttribute("r:id");
    const to = rId === null ? undefined : target.get(rId);
    if (to === undefined) problems.push(`the deck lists a slide whose relationship ${rId ?? "(none)"} does not exist`);
    else {
      const path = resolveTarget("ppt/presentation.xml", to ?? "");
      listed.add(path);
      if (!has(path)) problems.push(`the deck lists ${path}, which is not in the package`);
    }
  }
  // A slide part nothing lists is dead weight the user still carries.
  for (const name of names.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
    if (!listed.has(name)) problems.push(`${name} is in the package but not in the deck`);
  }
  return problems;
}

const DATA = `First\tLast\tShow
Ada\tLovelace\tyes
Grace\tHopper\t
Katherine\tJohnson\tyes`;

const CASES: { name: string; slides: SlideSpec[]; conditions?: Record<number, string> }[] = [
  { name: "one slide, three rows", slides: [{ paragraphs: [["Hello {{First}}"]] }] },
  {
    name: "a three-slide block",
    slides: [{ paragraphs: [["{{First}}"]] }, { paragraphs: [["{{Last}}"]] }, { paragraphs: [["no fields"]] }],
  },
  {
    // A row whose cell is empty drops its slide, so the plan is ragged and the
    // slide numbering has gaps to survive.
    name: "a conditional slide",
    slides: [{ paragraphs: [["{{First}}"]] }, { paragraphs: [["{{Last}}"]] }],
    conditions: { 2: "Show" },
  },
  {
    // A notes page is a second part per slide, with its own rels and its own
    // content-type override — the case `removeSlide` has to get right.
    name: "a slide with notes",
    slides: [{ paragraphs: [["{{First}}"]], notes: "note for {{Last}}" }],
  },
  {
    // A chart is FOUR parts per copy — the chart, its rels, its own workbook,
    // and the content-type override — and the workbook is a package inside the
    // package, declared by extension rather than by name. Every one of them is
    // a way to hand PowerPoint a file it calls damaged without saying which
    // part was wrong.
    name: "a slide with a chart and its workbook",
    slides: [
      {
        paragraphs: [["{{First}}"]],
        chart: { title: "Sales for {{Last}}", categories: ["{{First}}"], workbook: ["{{First}}"] },
      },
    ],
  },
  {
    // SmartArt relates to four parts and owns a fifth through one of them. The
    // drawing is the one a naive clone misses, and a copy pointing at a drawing
    // that is not there is a repair prompt.
    name: "a slide with SmartArt",
    slides: [{ paragraphs: [["{{First}}"]], smartArt: ["{{First}} and {{Last}}"] }],
  },
];

describe("the package the engine hands over", () => {
  for (const { name, slides, conditions } of CASES) {
    it(`is a legal package: ${name}`, async () => {
      // One slide past the block, so the deck is never only the template.
      const deck = await makeDeck([...slides, { paragraphs: [["after the block"]] }]);
      const pkg = await Pkg.open(deck);
      const prepared = await prepareBlock(
        pkg,
        { from: 1, to: slides.length, offsetInPackage: 0, ...(conditions ? { conditions } : {}) },
        "run1",
      );
      expect(prepared.ok, prepared.ok ? "" : prepared.why).toBe(true);
      if (!prepared.ok) return;
      const records = toRecordSet(parseDelimited(DATA));
      const plan = buildPlan(prepared.block, records, { runId: "run1" });
      const result = await runPlan(pkg, plan, records);
      expect(await problemsIn(await pkg.toBytes())).toEqual([]);

      // And the whole-deck route, which is what runs when the template came
      // back as the entire presentation: keep what the run produced, drop the
      // rest. That is where #38 lived, and it exercises `removeSlide` — the
      // one path that takes parts OUT of a package.
      const keep = new Set(result.slides);
      for (const path of await pkg.slidePaths()) if (!keep.has(path)) await pkg.removeSlide(path);
      expect(await pkg.slidePaths()).toEqual(result.slides);
      expect(await problemsIn(await pkg.toBytes())).toEqual([]);
    });
  }

  it("has something to say about a package that is broken", async () => {
    /**
     * The check that the check is not vacuous. Every assertion above is
     * `toEqual([])`, which is what an empty list of RULES also produces — and
     * this suite has twice caught a gate that measured nothing and reported
     * success. So a deliberately damaged package must come back with each kind
     * of complaint.
     */
    const deck = await makeDeck([{ paragraphs: [["a"]] }, { paragraphs: [["b"]] }]);
    const zip = await JSZip.loadAsync(deck);
    // Remove a slide part while leaving the deck listing it, and strip its
    // content type. One edit, three rules.
    zip.remove("ppt/slides/slide2.xml");
    const found = await problemsIn(await zip.generateAsync({ type: "uint8array" }));
    expect(found.some((p) => p.includes("not in the package"))).toBe(true);
    expect(found.some((p) => p.includes("content type declared for"))).toBe(true);

    // And the rule the generic one cannot reach: a slide typed only by the
    // package's `Default Extension="xml"`.
    const untyped = await JSZip.loadAsync(await makeDeck([{ paragraphs: [["a"]] }]));
    const types = await untyped.file("[Content_Types].xml")!.async("string");
    untyped.file("[Content_Types].xml", types.replace(/<Override PartName="\/ppt\/slides\/slide1\.xml"[^>]*\/>/, ""));
    const second = await problemsIn(await untyped.generateAsync({ type: "uint8array" }));
    expect(second).toContain("ppt/slides/slide1.xml has no content-type override of its own");
  });
});
