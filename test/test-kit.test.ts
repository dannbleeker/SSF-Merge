/**
 * The test kit's template, merged — against a chart THIS PROJECT DID NOT WRITE.
 *
 * Every other test here builds its own fixture, which means the engine is
 * checked against markup its own author wrote to the same understanding. A
 * chart is the case where that matters most: it has a string cache, a number
 * cache, a workbook of its own and half a dozen relationships, and a reader
 * built from the same misreading as the writer agrees with itself perfectly.
 *
 * `test-kit/SSF-Merge-test-template.pptx` was authored by python-pptx — a
 * different tool, a real chart part, a real embedded workbook — and committed.
 * Nothing in CI regenerates it, deliberately: the point is that it came from
 * somewhere else. `test-kit/build-template.py` records how, for the day it
 * needs rebuilding.
 *
 * The kit exists for a HUMAN to run in real PowerPoint (`docs/TEST-KIT.md`).
 * This test rides along on the same files so that the deck a person is asked to
 * open cannot quietly stop merging between rounds.
 *
 * It does NOT cover SmartArt: nothing outside PowerPoint can author one, so the
 * kit asks the tester to add it. `graphics.test.ts` covers SmartArt against
 * this project's own fixture, which is the weaker check and the only one
 * available.
 */
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { Pkg } from "../src/core/pptx/pkg.js";
import { prepareBlock } from "../src/core/merge/prepare.js";
import { buildPlan } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import { parseDelimited, toRecordSet } from "../src/core/data/recordset.js";
import { danglingRels } from "./fixtures/dangling.js";
import { A_NS, C_NS, CX_NS, MC_NS, PKG_REL_NS, SSML_NS, elements, parseXml } from "../src/core/pptx/xml.js";

const KIT = "test-kit";
const PHOTOS = ["ada.png", "grace.png", "alan.png"];

/** The whole kit run the way `docs/TEST-KIT.md` asks a person to run it. */
async function runTheKit() {
  const pkg = await Pkg.open(new Uint8Array(readFileSync(`${KIT}/SSF-Merge-test-template.pptx`)));
  // Slides 2 and 3, exactly as the instructions say. Slide 1 is the deck's own
  // instructions and is not part of the block.
  const prepared = await prepareBlock(pkg, { from: 2, to: 3, offsetInPackage: 1 }, "kit");
  if (!prepared.ok) throw new Error(`the kit's own template was refused: ${prepared.why}`);

  const records = toRecordSet(parseDelimited(readFileSync(`${KIT}/data.txt`, "utf8")));
  const images = new Map(PHOTOS.map((n) => [n, new Uint8Array(readFileSync(`${KIT}/${n}`))]));
  const plan = buildPlan(prepared.block, records, { runId: "kit" });
  const result = await runPlan(pkg, plan, records, { images });

  // The whole-deck route: keep what the run produced, drop the rest. That is
  // what the add-in hands PowerPoint, so it is what should be inspected.
  const keep = new Set(result.slides);
  for (const path of await pkg.slidePaths()) if (!keep.has(path)) await pkg.removeSlide(path);
  return { prepared, result, zip: await JSZip.loadAsync(await pkg.toBytes()) };
}

/** Resolve one part's relationships of a given type to package paths. */
async function related(zip: JSZip, part: string, endsWith: string): Promise<string[]> {
  const dir = part.slice(0, part.lastIndexOf("/"));
  const rels = await zip.file(`${dir}/_rels/${part.slice(part.lastIndexOf("/") + 1)}.rels`)?.async("string");
  if (!rels) return [];
  return elements(parseXml(rels), PKG_REL_NS, "Relationship")
    .filter((r) => (r.getAttribute("Type") ?? "").endsWith(endsWith))
    .map((r) => {
      const segments = dir.split("/");
      for (const seg of (r.getAttribute("Target") ?? "").split("/")) {
        if (seg === "..") segments.pop();
        else if (seg !== ".") segments.push(seg);
      }
      return segments.join("/");
    });
}

const kit = await runTheKit();
const slides = Object.keys(kit.zip.files)
  .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
  .sort((a, b) => Number(/(\d+)/.exec(a)?.[1]) - Number(/(\d+)/.exec(b)?.[1]));

describe("the kit's template, merged", () => {
  it("offers every field the instructions promise, including the one with no column", () => {
    // `Nickname` is in the template and NOT in the data, on purpose: a field
    // with no column stays visible on the slide, and the kit asks the tester to
    // confirm that. If this list changes, `docs/TEST-KIT.md` is wrong.
    expect(kit.prepared.ok && [...kit.prepared.fields].sort()).toEqual(
      ["Name", "Nickname", "Photo", "Region", "Renewal", "Revenue"].sort(),
    );
  });

  it("produces the six slides the instructions promise", () => {
    expect(kit.result.slides).toHaveLength(6);
    expect(slides).toHaveLength(6);
  });

  it("places one picture per record and reports nothing missing", () => {
    expect(kit.result.images).toMatchObject({ placed: 3, missing: [], unreadable: [] });
    const media = Object.keys(kit.zip.files).filter((n) => /^ppt\/media\/\w+\.png$/.test(n));
    expect(media, "one media part per distinct file").toHaveLength(3);
  });

  it("fills the chart a different tool wrote, in its title and in its category cache", async () => {
    const seen: string[] = [];
    for (const slide of slides) {
      const chart = (await related(kit.zip, slide, "/chart"))[0];
      if (!chart) continue;
      const doc = parseXml((await kit.zip.file(chart)?.async("string")) ?? "");
      const title = elements(doc, A_NS, "t")
        .map((t) => t.textContent ?? "")
        .join("");
      const cats = elements(doc, C_NS, "strCache").flatMap((c) =>
        elements(c, C_NS, "v").map((v) => v.textContent ?? ""),
      );
      expect(title).toMatch(/^Quarterly revenue — \w+$/);
      expect(cats, `${chart} kept a placeholder`).not.toContain("{{Region}}");
      seen.push(title);
    }
    // Three charts, three different regions: the copies are not sharing one.
    expect(seen).toEqual(["Quarterly revenue — Nordics", "Quarterly revenue — Benelux", "Quarterly revenue — DACH"]);
  });

  it("fills the workbook behind each chart, which is what Edit Data opens", async () => {
    expect(kit.result.graphics.workbooks).toBe(3);
    expect(kit.result.graphics.unreadable).toEqual([]);
    const regions: string[] = [];
    for (const slide of slides) {
      const chart = (await related(kit.zip, slide, "/chart"))[0];
      if (!chart) continue;
      const book = (await related(kit.zip, chart, "/package"))[0] ?? "";
      const inner = await JSZip.loadAsync((await kit.zip.file(book)?.async("uint8array")) as Uint8Array);
      const sst = (await inner.file("xl/sharedStrings.xml")?.async("string")) ?? "";
      const strings = elements(parseXml(sst), SSML_NS, "si").map((si) =>
        elements(si, SSML_NS, "t")
          .map((t) => t.textContent ?? "")
          .join(""),
      );
      expect(strings, `${book} kept a placeholder`).not.toContain("{{Region}}");
      regions.push(strings[0] ?? "");
    }
    expect(regions).toEqual(["Nordics", "Benelux", "DACH"]);
  });

  it("formats the number and the date the way the instructions say to check", async () => {
    const text = elements(parseXml((await kit.zip.file(slides[0] ?? "")?.async("string")) ?? ""), A_NS, "t").map(
      (t) => t.textContent ?? "",
    );
    expect(text).toContain("1 250 000 EUR");
    expect(text).toContain("Renewal 1 Mar 2026");
  });

  it("leaves the field with no column visible rather than blanking it", async () => {
    const text = elements(parseXml((await kit.zip.file(slides[1] ?? "")?.async("string")) ?? ""), A_NS, "t")
      .map((t) => t.textContent ?? "")
      .join(" ");
    expect(text).toContain("{{Nickname}}");
  });

  it("merges the speaker notes per copy", async () => {
    const notes = Object.keys(kit.zip.files).filter((n) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n));
    const said = await Promise.all(
      notes.map(async (n) =>
        elements(parseXml((await kit.zip.file(n)?.async("string")) ?? ""), A_NS, "t")
          .map((t) => t.textContent ?? "")
          .join(""),
      ),
    );
    expect(said.sort()).toEqual(["Call Ada before 1 Mar.", "Call Alan before 30 May.", "Call Grace before 15 Apr."]);
  });

  it("hands over a package with no relationship pointing at nothing", async () => {
    // The failure this whole round exists to catch, in the form a test can
    // reach: PowerPoint reports a package it cannot resolve as damaged, repairs
    // it silently, and drops whatever it chose to drop.
    expect(await danglingRels(kit.zip)).toEqual([]);
  });

  it("takes the template's own chart and workbook out with the template slides", () => {
    // Each copy has its own now, so the originals are referenced by nobody. A
    // whole chart and an embedded workbook per template slide is not a rounding
    // error in a package the host swallows as one base64 string.
    const parts = Object.keys(kit.zip.files);
    expect(parts.filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n))).toHaveLength(3);
    expect(parts.filter((n) => /^ppt\/embeddings\/\S+\.xlsx$/.test(n))).toHaveLength(3);
    expect(parts.some((n) => n.includes("Microsoft_Excel_Sheet1"))).toBe(false);
  });
});

/**
 * The second deck: a SUNBURST, written by real PowerPoint.
 *
 * `SSF-Merge-test-template.pptx` above holds a classic `<c:chartSpace>` chart
 * that python-pptx wrote. A modern chart is a different part under a different
 * relationship, and until this file arrived every chartEx test in the suite ran
 * against a fixture whose author was also the reader — the arrangement that let
 * a real defect through twice in one week, once in what a title looks like and
 * once in what `a:ext` means.
 *
 * Sunburst rather than another waterfall on purpose: a hierarchy chart keeps its
 * categories as SEVERAL `<cx:lvl>` inside one `<cx:strDim>`, so the level the
 * merge fills and the level it must leave alone are in the same element.
 *
 * `test-kit/strip-thinkcell.py` records the one edit made to this recording
 * before it was committed, and why.
 */
async function runTheSunburst() {
  const pkg = await Pkg.open(new Uint8Array(readFileSync(`${KIT}/modern-chart.pptx`)));
  const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "modern");
  if (!prepared.ok) throw new Error(`the kit's modern chart deck was refused: ${prepared.why}`);

  const records = toRecordSet(parseDelimited(readFileSync(`${KIT}/data.txt`, "utf8")));
  const result = await runPlan(pkg, buildPlan(prepared.block, records, { runId: "modern" }), records, {});

  const keep = new Set(result.slides);
  for (const path of await pkg.slidePaths()) if (!keep.has(path)) await pkg.removeSlide(path);
  return { prepared, result, zip: await JSZip.loadAsync(await pkg.toBytes()) };
}

const sun = await runTheSunburst();

/** One merged copy's chart part, in the order the copies were produced. */
async function chartsOf(zip: JSZip, slidePaths: string[]): Promise<Document[]> {
  const out: Document[] = [];
  for (const slide of slidePaths) {
    const part = (await related(zip, slide, "/chartEx"))[0];
    if (!part) continue;
    out.push(parseXml((await zip.file(part)?.async("string")) ?? ""));
  }
  return out;
}

/** The `<cx:pt>` text of one dimension, level by level. */
function levels(doc: Document, ns: string, local: string): string[][] {
  return elements(doc, ns, local).flatMap((dim) =>
    elements(dim, CX_NS, "lvl").map((lvl) => elements(lvl, CX_NS, "pt").map((pt) => pt.textContent ?? "")),
  );
}

describe("the kit's modern chart, merged", () => {
  it("reports the fields inside a chart PowerPoint wrote", () => {
    // `Name` is in three places — the slide's text box, the chart's title and
    // the series name — and `Region` in one cell of the category sheet. A block
    // whose only placeholder is inside a chart used to be refused as having
    // none.
    expect(sun.prepared.ok && [...sun.prepared.fields].sort()).toEqual(["Name", "Region"]);
  });

  it("gives every copy its own chart part and its own workbook", () => {
    const parts = Object.keys(sun.zip.files);
    expect(parts.filter((n) => /^ppt\/charts\/chartEx\d+\.xml$/.test(n))).toHaveLength(3);
    expect(parts.filter((n) => /^ppt\/embeddings\/\S+\.xlsx$/.test(n))).toHaveLength(3);
    // The template's own went out with the template slide.
    expect(parts.some((n) => n.includes("Microsoft_Excel_Worksheet"))).toBe(false);
  });

  it("fills the cell of the hierarchy that holds a placeholder, and no other", async () => {
    // The sheet is `Existing | {{Region}} | 100` over three rows: the INNER
    // level is a literal grouping and the outer one carries the placeholder in
    // its FIRST CELL ONLY. So a correct merge and a blind overwrite of every
    // category look different, which is the whole reason the deck is shaped
    // this way — Alan's outer level reads `DACH, Benelux, DACH`, and only the
    // first of those moved.
    //
    // Read what the grouping-level assertion can and cannot do. It catches a
    // merge that wrote where no placeholder was; it does NOT prove the engine
    // scopes `<cx:pt>` by its dimension, because nothing in this deck puts a
    // placeholder anywhere a scoping bug would reach. That guard needs a
    // placeholder in a value cell, which is a chart PowerPoint would not have
    // drawn, and it lives in `chart-modern.test.ts` where the fixture can hold
    // one.
    const outer: string[][] = [];
    for (const doc of await chartsOf(sun.zip, sun.result.slides)) {
      const [first, second] = levels(doc, CX_NS, "strDim");
      expect(second, "the grouping level was written to").toEqual(["Existing", "Existing", "New"]);
      outer.push(first ?? []);
    }
    expect(outer).toEqual([
      ["Nordics", "Benelux", "DACH"],
      ["Benelux", "Benelux", "DACH"],
      ["DACH", "Benelux", "DACH"],
    ]);
  });

  it("fills the chart's title, which PowerPoint keeps as DrawingML", async () => {
    // A modern chart's title can be `<a:t>` inside `<cx:rich>` or a bare
    // `<cx:v>`, and the engine fills both because two producers were seen doing
    // different things. This deck settles which one PowerPoint itself writes.
    const titles: string[] = [];
    for (const doc of await chartsOf(sun.zip, sun.result.slides)) {
      titles.push(
        elements(doc, A_NS, "t")
          .map((t) => t.textContent ?? "")
          .join(""),
      );
    }
    expect(titles).toEqual(["Ada pipeline", "Grace pipeline", "Alan pipeline"]);
  });

  it("fills the series name and leaves the plotted numbers alone", async () => {
    const names: string[] = [];
    for (const doc of await chartsOf(sun.zip, sun.result.slides)) {
      names.push(
        elements(doc, CX_NS, "txData")
          .flatMap((d) => elements(d, CX_NS, "v").map((v) => v.textContent ?? ""))
          .join(""),
      );
      // The plotted values, unchanged. Not the guard against writing a label
      // into a value — these numbers hold no placeholder, so a merge that
      // ignored the dimension entirely would still leave them alone, and this
      // line passes against that build. It pins that cloning a chart three ways
      // did not disturb what it plots; `chart-modern.test.ts` holds the scoping
      // guard, against a fixture with a placeholder in a value cell.
      expect(levels(doc, CX_NS, "numDim")).toEqual([["100", "60", "40"]]);
    }
    expect(names).toEqual(["Ada", "Grace", "Alan"]);
  });

  it("fills the workbook each chart opens on Edit Data", async () => {
    const seen: string[] = [];
    for (const slide of sun.result.slides) {
      const chart = (await related(sun.zip, slide, "/chartEx"))[0] ?? "";
      const book = (await related(sun.zip, chart, "/package"))[0] ?? "";
      const inner = await JSZip.loadAsync((await sun.zip.file(book)?.async("uint8array")) ?? new Uint8Array());
      const strings = parseXml((await inner.file("xl/sharedStrings.xml")?.async("string")) ?? "");
      const text = elements(strings, SSML_NS, "t").map((t) => t.textContent ?? "");
      expect(text, `${book} kept a placeholder`).not.toContain("{{Region}}");
      seen.push(text.join("|"));
    }
    // Closing Excel refreshes the chart from this, so a copy whose workbook
    // still said "{{Region}}" would undo its own merge on the first click.
    expect(seen).toEqual([
      "Ada|Existing|New|Nordics|Benelux|DACH",
      "Grace|Existing|New|Benelux|Benelux|DACH",
      "Alan|Existing|New|DACH|Benelux|DACH",
    ]);
  });

  it("swaps the fallback picture for the notice, on a slide PowerPoint laid out", async () => {
    for (const slide of sun.result.slides) {
      const doc = parseXml((await sun.zip.file(slide)?.async("string")) ?? "");
      const said = elements(doc, MC_NS, "Fallback").flatMap((f) =>
        elements(f, A_NS, "t").map((t) => t.textContent ?? ""),
      );
      expect(said.join(""), slide).toMatch(/needs a newer version of PowerPoint/);
      expect(elements(doc, MC_NS, "Fallback").flatMap((f) => elements(f, A_NS, "blip"))).toHaveLength(0);
      expect(await related(sun.zip, slide, "/image"), `${slide} still points at a rendering`).toEqual([]);
    }
  });

  it("hands over a package with no relationship pointing at nothing", async () => {
    expect(await danglingRels(sun.zip)).toEqual([]);
  });
});
