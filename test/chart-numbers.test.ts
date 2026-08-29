/**
 * A chart's NUMBERS, filled from the row.
 *
 * The text merge fills a chart's labels. The values could not be filled the
 * same way, because a `<c:numCache>` cell has to parse as a number and
 * `{{Revenue}}` does not — so the placeholder goes where the value actually
 * lives, in the embedded workbook's own cell, typed through Edit Data like any
 * other cell. No syntax of its own: the cell becomes an ordinary shared string.
 *
 * The thing worth testing hardest is that BOTH copies move. PowerPoint draws
 * from the chart's cache without opening the workbook; Excel opens the workbook
 * and refreshes the cache from it on close. Filling one and not the other is a
 * deck that looks right until somebody clicks Edit Data, which is exactly the
 * half-merge this project already knows from the label side.
 */
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { Pkg } from "../src/core/pptx/pkg.js";
import { prepareBlock } from "../src/core/merge/prepare.js";
import { buildPlan } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import { toRecordSet } from "../src/core/data/recordset.js";
import { cellsOfFormula } from "../src/core/merge/numbers.js";
import { makeDeck, type ChartSpec } from "./fixtures/deck.js";

const ROWS = [
  ["Name", "Revenue", "Notes"],
  ["Ada", "1250000", "hello"],
];

/** Merge one slide carrying this chart, and hand back the merged package. */
async function merge(chart: ChartSpec, rows = ROWS) {
  const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["{{Name}}"]], chart }]));
  const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "n");
  if (!prepared.ok) throw new Error(`the fixture was refused: ${prepared.why}`);
  const records = toRecordSet(rows);
  const plan = buildPlan(prepared.block, records, { runId: "n" });
  const out = await runPlan(pkg, plan, records);
  return { out, zip: await JSZip.loadAsync(await pkg.toBytes()) };
}

/** The merged chart's cached values, in point order. */
async function cachedValues(zip: JSZip): Promise<string[]> {
  const path = Object.keys(zip.files).find((n) => /^ppt\/charts\/chart2\.xml$/.test(n));
  const xml = await zip.file(path!)!.async("string");
  const cache = /<c:numCache>[\s\S]*?<\/c:numCache>/.exec(xml)?.[0] ?? "";
  return [...cache.matchAll(/<c:v>([^<]*)<\/c:v>/g)].map((m) => m[1]!);
}

/**
 * The MERGED chart's own workbook, resolved through its relationships.
 *
 * Not "the first embedding in the package": the merged deck still holds the
 * template's chart and its untouched workbook, so taking any one of them read
 * the wrong file and reported the feature broken while it worked. Which
 * workbook belongs to which chart is the whole point of the pairing.
 */
async function sheetCells(zip: JSZip): Promise<Record<string, { type: string; value: string }>> {
  const rels = await zip.file("ppt/charts/_rels/chart2.xml.rels")!.async("string");
  const target = /Target="([^"]*\.xlsx)"/.exec(rels)?.[1] ?? "";
  const path = `ppt/${target.replace(/^\.\.\//, "")}`;
  const book = await JSZip.loadAsync(await zip.file(path)!.async("nodebuffer"));
  const sheet = await book.file("xl/worksheets/sheet1.xml")!.async("string");
  const sst = (await book.file("xl/sharedStrings.xml")?.async("string")) ?? "";
  const strings = [...sst.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) =>
    [...m[1]!.matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map((t) => t[1]).join(""),
  );
  const out: Record<string, { type: string; value: string }> = {};
  for (const m of sheet.matchAll(/<c r="([A-Z]+\d+)"([^>]*)>([\s\S]*?)<\/c>/g)) {
    const type = /t="([a-z]+)"/.exec(m[2]!)?.[1] ?? "n";
    const raw = /<v>([^<]*)<\/v>/.exec(m[3]!)?.[1] ?? "";
    out[m[1]!] = { type, value: type === "s" ? (strings[Number(raw)] ?? "") : raw };
  }
  return out;
}

describe("reading which cell a cached point came from", () => {
  it("walks a column, a row and a single cell", () => {
    expect(cellsOfFormula("Sheet1!$B$2:$B$4")).toEqual(["B2", "B3", "B4"]);
    expect(cellsOfFormula("Sheet1!$B$2:$D$2")).toEqual(["B2", "C2", "D2"]);
    expect(cellsOfFormula("Sheet1!$B$1")).toEqual(["B1"]);
    expect(cellsOfFormula("'My Sheet'!$B$2:$B$3")).toEqual(["B2", "B3"]);
  });

  it("refuses a rectangle rather than guessing its order", () => {
    // A series does not read one, and walking it the wrong way would pair a
    // value with the wrong point — which no count would catch.
    expect(cellsOfFormula("Sheet1!$B$2:$D$4")).toBeNull();
    expect(cellsOfFormula("nonsense")).toBeNull();
  });
});

describe("a value cell holding a placeholder", () => {
  const spec: ChartSpec = {
    categories: ["{{Name}}", "Everyone else"],
    workbook: ["{{Name}}", "Everyone else"],
    values: ["{{Revenue}}", "42"],
  };

  it("fills the chart's cached number", async () => {
    const { zip } = await merge(spec);
    expect(await cachedValues(zip)).toEqual(["1250000", "42"]);
  });

  it("fills the workbook cell too, and makes it a NUMBER again", async () => {
    // The half that Excel reads. Left as a shared string it would show the
    // number as text, plot nothing, and revert the cache on close.
    const { zip } = await merge(spec);
    const cells = await sheetCells(zip);
    expect(cells["B2"]).toEqual({ type: "n", value: "1250000" });
  });

  it("agrees with itself in both places", async () => {
    const { zip } = await merge(spec);
    const cached = await cachedValues(zip);
    const cells = await sheetCells(zip);
    expect(cells["B2"]!.value, "the cache and the workbook disagree").toBe(cached[0]);
  });

  it("counts what it filled", async () => {
    const { out } = await merge(spec);
    expect(out.graphics.numbers.filled).toBe(1);
    expect(out.graphics.numbers.refused).toBe(0);
  });
});

describe("what it leaves alone", () => {
  it("an ordinary number", async () => {
    const { zip } = await merge({
      workbook: ["{{Name}}"],
      values: ["{{Revenue}}", "42"],
      categories: ["{{Name}}", "Everyone else"],
    });
    expect((await sheetCells(zip))["B3"]).toEqual({ type: "n", value: "42" });
  });

  it("a value that will not be a number, rather than writing a zero", async () => {
    // Guessing zero would draw a chart the data never said. The placeholder
    // stays exactly as written, the same answer a column-less placeholder gets
    // on a slide.
    const { out, zip } = await merge({
      workbook: ["{{Name}}"],
      values: ["{{Notes}}", "42"],
      categories: ["{{Name}}", "Everyone else"],
    });
    expect(out.graphics.numbers.refused).toBe(1);
    expect(out.graphics.numbers.filled).toBe(0);
    const cells = await sheetCells(zip);
    expect(cells["B2"]!.type, "a refused cell was rewritten anyway").toBe("s");
  });
});

describe("the number that reaches the chart", () => {
  it("is the raw value, not the formatted one", async () => {
    // `{{Revenue|number:0}}` in a LABEL should read "1 250 000". In a value cell
    // that string plots nothing, so the format is deliberately not applied —
    // a chart formats its own axis.
    const { zip } = await merge({
      workbook: ["{{Name}}"],
      values: ["{{Revenue|number:0}}", "42"],
      categories: ["{{Name}}", "Everyone else"],
    });
    expect((await sheetCells(zip))["B2"]).toEqual({ type: "n", value: "1250000" });
  });

  it("reads the European form a spreadsheet exports", async () => {
    const { zip } = await merge(
      { workbook: ["{{Name}}"], values: ["{{Revenue}}", "42"], categories: ["{{Name}}", "Everyone else"] },
      [
        ["Name", "Revenue"],
        ["Ada", "1.250.000,50"],
      ],
    );
    expect((await sheetCells(zip))["B2"]!.value).toBe("1250000.5");
  });
});
