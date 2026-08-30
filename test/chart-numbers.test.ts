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
import { cellsOfFormula, modernSeries, sheetOfFormula } from "../src/core/merge/numbers.js";
import { parseXml } from "../src/core/pptx/xml.js";
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

describe("a value cell that is not ONLY a placeholder", () => {
  /**
   * The cell is rebuilt around its placeholders rather than handed to
   * `String.replace`, since the reader stopped being a regular expression. That
   * is index arithmetic, and index arithmetic that nothing checks is how a
   * merge writes a number nobody asked for — so the characters before the first
   * placeholder, between two, and after the last are each pinned by a case that
   * fails without them.
   *
   * None of these is a template anybody would write. They are the shapes the
   * rebuild can get wrong.
   */
  it("keeps what comes after the placeholder", async () => {
    const { zip } = await merge({
      workbook: ["{{Name}}"],
      values: ["{{Revenue}}00", "42"],
      categories: ["{{Name}}", "Everyone else"],
    });
    expect((await sheetCells(zip))["B2"]).toEqual({ type: "n", value: "125000000" });
  });

  it("keeps what comes before it", async () => {
    const { zip } = await merge({
      workbook: ["{{Name}}"],
      values: ["-{{Revenue}}", "42"],
      categories: ["{{Name}}", "Everyone else"],
    });
    expect((await sheetCells(zip))["B2"]).toEqual({ type: "n", value: "-1250000" });
  });

  it("keeps what sits between two of them", async () => {
    const { zip } = await merge(
      {
        workbook: ["{{Name}}"],
        values: ["{{A}}.{{B}}", "42"],
        categories: ["{{Name}}", "Everyone else"],
      },
      [
        ["Name", "A", "B"],
        ["Ada", "3", "5"],
      ],
    );
    expect((await sheetCells(zip))["B2"]).toEqual({ type: "n", value: "3.5" });
  });

  it("leaves a placeholder no column answers standing, braces and all", async () => {
    // The cell then fails to parse and is REFUSED, which is the point: blanking
    // an unresolved placeholder would leave `42` behind and plot it as though
    // the data had said so.
    const { out, zip } = await merge({
      workbook: ["{{Name}}"],
      values: ["{{Nope}}42", "42"],
      categories: ["{{Name}}", "Everyone else"],
    });
    expect(out.graphics.numbers.refused).toBe(1);
    expect(out.graphics.numbers.filled).toBe(0);
    expect((await sheetCells(zip))["B2"]!.type, "a refused cell was rewritten anyway").toBe("s");
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

describe("which sheet the formula names", () => {
  it("reads the sheet title out of the formula", () => {
    expect(sheetOfFormula("Sheet1!$B$2:$B$3")).toBe("Sheet1");
    expect(sheetOfFormula("'My Sheet'!$B$2")).toBe("My Sheet");
    expect(sheetOfFormula("'It''s Data'!$B$2")).toBe("It's Data");
    expect(sheetOfFormula("$B$2")).toBeNull();
  });

  it("writes into the sheet the chart points at, not the first one holding the address", async () => {
    /**
     * The first version of this pass searched every sheet for the address and
     * took the first hit, reasoning that an embedded chart workbook has one
     * sheet. True of every one it had met, and not a fact about the format —
     * add a sheet in Edit Data and `B2` exists twice.
     *
     * The failure it produced would have been silent and invisible to a count:
     * the right NUMBER of cells is written, just not the ones the chart plots.
     * So this fixture puts a DECOY `B2` on the sheet that sorts first and
     * points the chart at the second.
     */
    const pkg = await Pkg.open(
      await makeDeck([
        {
          paragraphs: [["{{Name}}"]],
          chart: { categories: ["{{Name}}", "Other"], workbook: ["{{Name}}"], values: ["{{Revenue}}", "42"] },
        },
      ]),
    );

    const S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
    const REL = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
    const embedding = pkg.partNames().find((n) => /^ppt\/embeddings\/.+\.xlsx$/.test(n))!;
    const book = await JSZip.loadAsync(await pkg.bytes(embedding));

    // sheet1 becomes the decoy, holding a plain number at B2.
    book.file(
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet ${S}><sheetData>` +
        `<row r="2"><c r="B2"><v>999</v></c></row></sheetData></worksheet>`,
    );
    // sheet2 is the real one: B2 is the placeholder, as a shared string.
    const sst = await book.file("xl/sharedStrings.xml")!.async("string");
    const count = (sst.match(/<si>/g) ?? []).length;
    book.file("xl/sharedStrings.xml", sst.replace("</sst>", `<si><t>{{Revenue}}</t></si></sst>`));
    book.file(
      "xl/worksheets/sheet2.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet ${S}><sheetData>` +
        `<row r="2"><c r="B2" t="s"><v>${count}</v></c></row></sheetData></worksheet>`,
    );
    book.file(
      "xl/workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<workbook ${S} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>` +
        `<sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId3"/>` +
        `</sheets></workbook>`,
    );
    book.file(
      "xl/_rels/workbook.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>` +
        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` +
        `</Relationships>`,
    );
    pkg.setBytes(embedding, await book.generateAsync({ type: "uint8array" }));
    pkg.setText(
      "ppt/charts/chart1.xml",
      (await pkg.text("ppt/charts/chart1.xml")).replace("Sheet1!$B$2:$B$3", "Sheet2!$B$2:$B$3"),
    );

    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "s");
    if (!prepared.ok) throw new Error(prepared.why);
    const records = toRecordSet(ROWS);
    await runPlan(pkg, buildPlan(prepared.block, records, { runId: "s" }), records);

    const zip = await JSZip.loadAsync(await pkg.toBytes());
    const rels = await zip.file("ppt/charts/_rels/chart2.xml.rels")!.async("string");
    const merged = await JSZip.loadAsync(
      await zip.file(`ppt/${/Target="([^"]*\.xlsx)"/.exec(rels)![1]!.replace(/^\.\.\//, "")}`)!.async("nodebuffer"),
    );
    const two = await merged.file("xl/worksheets/sheet2.xml")!.async("string");
    const one = await merged.file("xl/worksheets/sheet1.xml")!.async("string");

    expect(two, "the sheet the chart names did not get the value").toContain("<v>1250000</v>");
    expect(one, "the decoy sheet was written into instead").toContain("<v>999</v>");
    expect(one).not.toContain("1250000");
  });
});

describe("a workbook whose worksheet part is not named sheet1.xml", () => {
  /**
   * The worksheet list was collected by matching part names against
   * `xl/worksheets/sheetN.xml`. Excel writes that name; a workbook built by
   * anything else need not, and the format does not require it.
   *
   * The result was the quietest failure in this engine: not a wrong number, not
   * a refusal, but NOTHING. `filled: 0`, `refused: 0`, the chart keeping its
   * cached values, the deck looking finished — and the placeholder still
   * sitting in the cell for whoever eventually clicks Edit Data.
   *
   * The workbook says which parts are its sheets, in `xl/workbook.xml` and its
   * relationships. That is read now instead of the file names.
   */
  const spec: ChartSpec = {
    categories: ["{{Name}}", "Everyone else"],
    workbook: ["{{Name}}", "Everyone else"],
    values: ["{{Revenue}}", "42"],
  };

  /** Rename the worksheet part, keeping every declaration that names it true. */
  async function renameWorksheet(pkg: Pkg, to: string): Promise<void> {
    const from = "xl/worksheets/sheet1.xml";
    const embedding = pkg.partNames().find((p) => p.endsWith(".xlsx"));
    const book = await JSZip.loadAsync(await pkg.bytes(embedding as string));
    book.file(to, await book.file(from)!.async("string"));
    book.remove(from);
    const rels = await book.file("xl/_rels/workbook.xml.rels")!.async("string");
    book.file("xl/_rels/workbook.xml.rels", rels.replace("worksheets/sheet1.xml", to.replace("xl/", "")));
    const types = await book.file("[Content_Types].xml")!.async("string");
    book.file("[Content_Types].xml", types.replace(`/${from}`, `/${to}`));
    pkg.setBytes(embedding as string, await book.generateAsync({ type: "uint8array" }));
  }

  async function mergeRenamed(to: string) {
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["{{Name}}"]], chart: spec }]));
    await renameWorksheet(pkg, to);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "n");
    if (!prepared.ok) throw new Error(`the fixture was refused: ${prepared.why}`);
    const records = toRecordSet(ROWS);
    return runPlan(pkg, buildPlan(prepared.block, records, { runId: "n" }), records);
  }

  it("fills its numbers all the same", async () => {
    const out = await mergeRenamed("xl/worksheets/data.xml");
    expect(out.graphics.numbers, "the sheet was found by its file name, not its declaration").toEqual({
      filled: 1,
      refused: 0,
    });
  });

  it("fills them when the part is not under xl/worksheets at all", async () => {
    // The relationship target decides the path. Nothing in the format says a
    // worksheet lives in that folder either.
    const out = await mergeRenamed("xl/data.xml");
    expect(out.graphics.numbers).toEqual({ filled: 1, refused: 0 });
  });
});

describe("which cached numbers a MODERN chart offers", () => {
  // Two shapes the format allows and nothing here can author, so they are
  // driven directly rather than through a deck. Both are refusals, and a
  // refusal no test can reach is one nobody can check.
  const dim = (inner: string) =>
    parseXml(
      `<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">` +
        `<cx:numDim type="val">${inner}</cx:numDim></cx:chartSpace>`,
    );

  it("reads the one level a dimension normally has", () => {
    const series = modernSeries(
      dim(`<cx:f>Sheet1!$B$2:$B$3</cx:f><cx:lvl ptCount="2"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt></cx:lvl>`),
    );
    expect(series).toHaveLength(1);
    expect(series[0]?.formula).toBe("Sheet1!$B$2:$B$3");
    expect(series[0]?.points.map((p) => p.idx)).toEqual([0, 1]);
  });

  it("ignores an empty level beside a full one", () => {
    // A sunburst's category dimension carries a trailing `<cx:lvl ptCount="0"/>`,
    // and a numeric one may too. Counting it would make the dimension look
    // multi-level and refuse a perfectly ordinary series.
    const series = modernSeries(
      dim(`<cx:f>Sheet1!$B$2</cx:f><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl><cx:lvl ptCount="0"/>`),
    );
    expect(series).toHaveLength(1);
  });

  it("refuses a dimension with more than one populated level", () => {
    // Its range is a rectangle, and which level is which column is a guess.
    // `cellsOfFormula` would refuse the rectangle anyway; this does not rely on
    // that, because "the other refusal happens to cover it" stops being true
    // the moment somebody writes a multi-level dimension over a single column.
    const series = modernSeries(
      dim(
        `<cx:f>Sheet1!$A$2:$B$3</cx:f>` +
          `<cx:lvl ptCount="2"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt></cx:lvl>` +
          `<cx:lvl ptCount="2"><cx:pt idx="0">3</cx:pt><cx:pt idx="1">4</cx:pt></cx:lvl>`,
      ),
    );
    expect(series).toEqual([]);
  });

  it("offers a dimension with no formula, and the empty range refuses it", () => {
    // `<cx:f>` is optional: a dimension may carry literal data with no workbook
    // behind it. There is then no cell to fill, and the empty formula is
    // refused by the same reader that refuses a range it cannot parse.
    const series = modernSeries(dim(`<cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl>`));
    expect(series).toHaveLength(1);
    expect(series[0]?.formula).toBe("");
    expect(cellsOfFormula(series[0]?.formula ?? "")).toBeNull();
  });
});
