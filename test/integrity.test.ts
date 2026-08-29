/**
 * Whether the packages this engine produces are internally consistent.
 *
 * A different question from every other test here. They ask whether a merged
 * deck SAYS the right things — this row's region in this row's chart — and a
 * deck can say all of it correctly while naming a relationship that is not
 * there, which PowerPoint calls damage and repairs by dropping whatever it
 * chooses.
 *
 * It exists because two defects found on 2026-08-29 were the same mistake in
 * the two places that delete a relationship, and no gate here could see either.
 * The rule those defects broke is not about charts or tags: it is that the
 * parts, the relationships and the markup have to agree. So the check is
 * structural, knows nothing about what a deck is for, and is pointed at
 * everything this suite can build.
 *
 * `scripts/package-integrity.mjs` holds the rules, so the human round's
 * verifier and this file cannot read different ones.
 */
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import JSZip from "jszip";
// @ts-expect-error — a plain .mjs tool with no types. The rules live THERE so
// the suite and `test-kit/driver/verify-package.mjs` share one copy.
import { packageProblems } from "../scripts/package-integrity.mjs";
import { Pkg } from "../src/core/pptx/pkg.js";
import { prepareBlock } from "../src/core/merge/prepare.js";
import { buildPlan } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import { parseDelimited, toRecordSet } from "../src/core/data/recordset.js";
import { A_NS, elements } from "../src/core/pptx/xml.js";
import { makeDeck, type SlideSpec } from "./fixtures/deck.js";

type Parts = Map<string, string | Uint8Array>;

async function partsOf(bytes: Uint8Array): Promise<Parts> {
  const zip = await JSZip.loadAsync(bytes);
  const parts: Parts = new Map();
  for (const name of Object.keys(zip.files)) {
    const file = zip.file(name);
    if (!file || file.dir) continue;
    const xml = name.endsWith(".xml") || name.endsWith(".rels");
    parts.set(name, xml ? await file.async("string") : await file.async("uint8array"));
  }
  return parts;
}

const problems = (parts: Parts): string[] => packageProblems(parts) as string[];

/** Merge a block and hand back the package, optionally minus the template. */
async function merged(bytes: Uint8Array, from: number, to: number, sweep: boolean): Promise<Uint8Array> {
  const pkg = await Pkg.open(bytes);
  const prepared = await prepareBlock(pkg, { from, to, offsetInPackage: from - 1 }, "i");
  if (!prepared.ok) throw new Error(`refused: ${prepared.why}`);
  const records = toRecordSet(parseDelimited(readFileSync("test-kit/data.txt", "utf8")));
  const images = new Map(
    ["ada.png", "grace.png", "alan.png"].map((n) => [n, new Uint8Array(readFileSync(`test-kit/${n}`))]),
  );
  const result = await runPlan(pkg, buildPlan(prepared.block, records, { runId: "i" }), records, { images });
  if (sweep) {
    const keep = new Set(result.slides);
    for (const path of await pkg.slidePaths()) if (!keep.has(path)) await pkg.removeSlide(path);
  }
  return pkg.toBytes();
}

/** Everything the fixture can put on one slide, which no other test combines. */
const EVERYTHING: SlideSpec = {
  paragraphs: [["{{Name}}"]],
  notes: "note {{Name}}",
  chart: { title: "{{Name}}", categories: ["{{Name}}", "b"], workbook: ["{{Name}}"], values: ["1", "2"] },
  modernChart: { title: "{{Name}}", categories: ["{{Name}}"], series: "{{Name}}", workbook: ["{{Name}}"] },
  smartArt: ["{{Name}}", "second"],
  smartArtDrawingOn: "slide",
  icons: true,
  shapeTags: true,
};

describe("the packages this engine hands over", () => {
  it("finds nothing wrong with the decks as they are committed", async () => {
    // Both are recordings, and a problem here would mean the recording changed
    // rather than the engine — worth knowing which, so they are checked apart
    // from anything the merge did to them.
    for (const path of ["test-kit/SSF-Merge-test-template.pptx", "test-kit/modern-chart.pptx"]) {
      expect(problems(await partsOf(new Uint8Array(readFileSync(path)))), path).toEqual([]);
    }
  });

  it("finds nothing wrong with the kit merged, the way the manual asks", async () => {
    const bytes = await merged(new Uint8Array(readFileSync("test-kit/SSF-Merge-test-template.pptx")), 2, 3, true);
    expect(problems(await partsOf(bytes))).toEqual([]);
  });

  it("finds nothing wrong with a modern chart PowerPoint wrote, merged", async () => {
    const bytes = await merged(new Uint8Array(readFileSync("test-kit/modern-chart.pptx")), 1, 1, true);
    expect(problems(await partsOf(bytes))).toEqual([]);
  });

  it("finds nothing wrong with every feature at once, template kept or swept", async () => {
    // The combination is the point. A slide carrying a chart AND a modern chart
    // AND SmartArt AND an icon AND a shape's own tags is where the passes that
    // rewrite relationships meet each other, and until this existed no test put
    // them together — which is how the fixture came to describe a slide whose
    // diagram colours and modern chart shared a relationship id.
    for (const sweep of [false, true]) {
      const deck = await makeDeck([EVERYTHING, { paragraphs: [["after"]] }]);
      const bytes = await merged(deck, 1, 1, sweep);
      expect(problems(await partsOf(bytes)), sweep ? "swept" : "template kept").toEqual([]);
    }
  });

  it("finds nothing wrong with a block whose slides are all different", async () => {
    // Every other case here is one template slide repeated. A block of three,
    // each carrying different parts, is what exercises the per-PACKAGE counters
    // — tag numbers, notes numbering, chart numbering — against each other
    // across one run, and nothing did.
    //
    // Also worth having as an ordering check: the plan is record-major, so the
    // nine slides read A/B/C for Ada, then A/B/C for Bo. A counter that handed
    // two copies the same part number would show up here as a slide carrying
    // somebody else's chart.
    const deck = await makeDeck([
      { paragraphs: [["A {{Name}}"]], notes: "notes A {{Name}}", shapeTags: true },
      {
        paragraphs: [["B {{Name}}"]],
        chart: { title: "{{Name}}", categories: ["{{Name}}"], workbook: ["{{Name}}"], values: ["1"] },
        icons: true,
      },
      {
        paragraphs: [["C {{Name}}"]],
        modernChart: { title: "{{Name}}", categories: ["{{Name}}"], series: "{{Name}}", workbook: ["{{Name}}"] },
        smartArt: ["{{Name}}", "b"],
        smartArtDrawingOn: "slide",
        notes: "notes C {{Name}}",
      },
      { paragraphs: [["after"]] },
    ]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 3, offsetInPackage: 0 }, "multi");
    if (!prepared.ok) throw new Error(`refused: ${prepared.why}`);
    const records = toRecordSet([["Name"], ["Ada"], ["Bo"], ["Cy"]]);
    const result = await runPlan(pkg, buildPlan(prepared.block, records, { runId: "multi" }), records, {});
    expect(result.slides).toHaveLength(9);

    const keep = new Set(result.slides);
    for (const path of await pkg.slidePaths()) if (!keep.has(path)) await pkg.removeSlide(path);
    const bytes = await pkg.toBytes();
    expect(problems(await partsOf(bytes))).toEqual([]);

    // Record-major, and each copy carrying its own template slide's text.
    const said: string[] = [];
    for (const slide of result.slides) {
      const doc = await pkg.doc(slide);
      said.push((elements(doc, A_NS, "t")[0]?.textContent ?? "").trim());
    }
    expect(said).toEqual(["A Ada", "B Ada", "C Ada", "A Bo", "B Bo", "C Bo", "A Cy", "B Cy", "C Cy"]);
  });
});

describe("the checker itself", () => {
  /**
   * A checker that has never failed is a checker nobody has tested.
   *
   * Each case breaks a good package one way and asserts the problem is named.
   * The four are the four the engine could produce, and two of them are exactly
   * what #124 and #126 shipped.
   */
  async function goodParts(): Promise<Parts> {
    const deck = await makeDeck([EVERYTHING, { paragraphs: [["after"]] }]);
    return partsOf(await merged(deck, 1, 1, true));
  }

  const slideOf = (parts: Parts): string =>
    [...parts.keys()].filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort()[0] ?? "";

  it("names a relationship whose target is not in the package", async () => {
    const parts = await goodParts();
    const slide = slideOf(parts);
    const rels = `ppt/slides/_rels/${slide.split("/").pop()}.rels`;
    parts.set(rels, (parts.get(rels) as string).replace(/Target="\.\.\/charts\/[^"]*"/, 'Target="../charts/gone.xml"'));
    expect(problems(parts).join("\n")).toContain("which is not in the package");
  });

  it("names markup that references a relationship the part does not have", async () => {
    // #124: the relationship deleted out from under live markup.
    const parts = await goodParts();
    const slide = slideOf(parts);
    const rels = `ppt/slides/_rels/${slide.split("/").pop()}.rels`;
    const body = parts.get(rels) as string;
    const id = /Id="([^"]*)"[^>]*relationships\/image/.exec(body)?.[1] ?? "";
    expect(id, "the fixture drew no image relationship, so this proves nothing").not.toBe("");
    parts.set(rels, body.replace(new RegExp(`<Relationship Id="${id}"[^>]*/>`), ""));
    expect(problems(parts).join("\n")).toContain("names a relationship the part does not have");
  });

  it("names a reference that leads to the wrong KIND of part", async () => {
    // #126: the id freed by a delete and taken by the next thing that needed
    // one, so the reference still resolves — to somebody else's data. Both
    // checks above pass on this.
    const parts = await goodParts();
    const slide = slideOf(parts);
    const rels = `ppt/slides/_rels/${slide.split("/").pop()}.rels`;
    parts.set(
      rels,
      (parts.get(rels) as string).replace(
        /(<Relationship Id="[^"]*" Type="[^"]*)relationships\/tags("[^>]*Target=")[^"]*"/,
        '$1relationships/image$2../media/whatever.png"',
      ),
    );
    const said = problems(parts).join("\n");
    expect(said).toContain("wants tags and leads to image");
  });

  it("names a part no content type covers", async () => {
    // An EXTENSION with no Default, which is the shape this can actually take.
    // A first attempt added a stray `.xml` part and the check said nothing —
    // rightly: every deck declares `Default Extension="xml"`, so that one is
    // covered by definition. The real case is an embedding or a media file
    // whose extension nothing declares, which is what stripping a part out of a
    // package leaves behind if its Default goes with it.
    const parts = await goodParts();
    parts.set("ppt/embeddings/oleObject1.bin", new Uint8Array([1, 2, 3]));
    expect(problems(parts).join("\n")).toContain("no content type covers it");
  });
});
