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
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { Pkg } from "../src/core/pptx/pkg.js";
import { prepareBlock } from "../src/core/merge/prepare.js";
import { buildPlan } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import { toRecordSet } from "../src/core/data/recordset.js";
import { cellAt, cellsOfFormula, MAX_SERIES_CELLS, modernSeries, sheetOfFormula } from "../src/core/merge/numbers.js";
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

  it("refuses a range too long to be a series, rather than allocating until the process dies", () => {
    /**
     * `Sheet1!A1:A99999999` killed the process — not a hang and not an
     * exception the pane could show, but a fatal out-of-memory abort with no
     * JS stack. It is reached from `prepareBlock`, which runs when the user
     * picks the template block, before any data has been pasted.
     *
     * The endpoint pattern bounds the SHAPE of an address and not its
     * magnitude, so the one range shape this function accepts was the one with
     * no ceiling. `A1:ZZZZZZ1` is the same defect along the column axis.
     */
    expect(cellsOfFormula("Sheet1!A1:A99999999")).toBeNull();
    expect(cellsOfFormula("Sheet1!A1:ZZZZZZ1")).toBeNull();
    // Excel's own full-column reference, which it writes whenever somebody
    // selects a whole column as a chart series. A legal deck, and it was inside
    // no bound at all.
    expect(cellsOfFormula("Sheet1!$A$1:$A$1048576")).toBeNull();
  });

  it("takes a series right up to the bound and refuses one cell past it", () => {
    // The boundary itself, so the bound cannot be quietly widened or narrowed
    // without this saying so.
    expect(cellsOfFormula(`Sheet1!A1:A${MAX_SERIES_CELLS}`)).toHaveLength(MAX_SERIES_CELLS);
    expect(cellsOfFormula(`Sheet1!A1:A${MAX_SERIES_CELLS + 1}`)).toBeNull();
  });

  it("refuses a rectangle rather than guessing its order", () => {
    // A series does not read one, and walking it the wrong way would pair a
    // value with the wrong point — which no count would catch.
    expect(cellsOfFormula("Sheet1!$B$2:$D$4")).toBeNull();
    expect(cellsOfFormula("nonsense")).toBeNull();
  });
});

describe("finding a cell in a sheet", () => {
  /** A worksheet holding `n` cells down column A. */
  function sheetOf(n: number): Document {
    const cells = Array.from({ length: n }, (_, i) => `<c r="A${i + 1}"><v>${i}</v></c>`).join("");
    return parseXml(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row>${cells}</row></sheetData></worksheet>`,
    );
  }

  it("walks the sheet ONCE however many cells are asked for", () => {
    /**
     * It walked the whole worksheet for each address, so reading a series was
     * quadratic in the sheet. A legal `Sheet1!$A$1:$A$1048576` over twenty rows
     * of data took 67 seconds against 194 ms for the same chart with a two-cell
     * range, and a 240-row merge of it extrapolates past thirteen minutes — in
     * a task-pane WebView, with nothing on screen to say why.
     *
     * The assertion is on WORK rather than on wall clock: a stopwatch in a
     * suite measures the machine it happens to run on, and this measures the
     * thing that made it slow.
     */
    const sheet = sheetOf(500);
    let walks = 0;
    const real = sheet.getElementsByTagNameNS.bind(sheet);
    sheet.getElementsByTagNameNS = ((ns: string, local: string) => {
      walks++;
      return real(ns, local);
    }) as typeof sheet.getElementsByTagNameNS;

    for (let i = 1; i <= 500; i++) expect(cellAt(sheet, `A${i}`)?.getAttribute("r")).toBe(`A${i}`);
    expect(walks, "one walk per lookup is what made a legal chart take minutes").toBe(1);
    // A miss is answered from the same index, not by walking again.
    expect(cellAt(sheet, "ZZ99")).toBeUndefined();
    expect(walks).toBe(1);
  });

  it("is sound only while nothing adds a cell to a sheet", () => {
    /**
     * The index lives as long as the sheet document, so a `<c>` created after a
     * lookup would be invisible to it and the merge would leave that cell as
     * the author typed it, with nothing said. The one write this pass makes is
     * a `<v>` inside a cell that already exists.
     *
     * A source scan rather than a behavioural test, because the defect it
     * guards against is a call site that does not exist yet — nothing can
     * observe an absence behaviourally.
     */
    const source = readFileSync("src/core/merge/numbers.ts", "utf8");
    const created = [...source.matchAll(/createElementNS\(\s*SSML_NS\s*,\s*"([^"]+)"/g)].map((m) => m[1]);
    expect(created, "a new cell would be invisible to the index above").not.toContain("c");
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

/**
 * A value cell that does not go through the shared-string table.
 *
 * Excel writes `t="s"` and an index; a generator that never built a
 * shared-string table writes the text inline instead, and `mergeGraphics`
 * already reads worksheets for exactly that reason — "a chart written by a tool
 * rather than by Excel". The branch that reads one had never been exercised:
 * every fixture here goes through the shared table, so coverage showed the
 * inline arm at zero on the one pass whose output a user opens in Excel.
 */
describe("a value cell that holds its string inline", () => {
  /** Rewrite the fixture's B2 value cell to the given raw XML. */
  async function withValueCell(cellXml: string) {
    const deck = await makeDeck([
      { paragraphs: [["{{Name}}"]], chart: { categories: ["a"], workbook: ["x"], values: ["0"] } },
    ]);
    const zip = await JSZip.loadAsync(deck);
    const emb = Object.keys(zip.files).find((n) => /embeddings\/.*\.xlsx$/.test(n))!;
    const book = await JSZip.loadAsync(await zip.file(emb)!.async("nodebuffer"));
    const sheet = await book.file("xl/worksheets/sheet1.xml")!.async("string");
    const patched = sheet.replace('<c r="B2"><v>0</v></c>', cellXml);
    expect(patched, "the fixture's value cell moved; this test patches it by hand").not.toBe(sheet);
    book.file("xl/worksheets/sheet1.xml", patched);
    zip.file(emb, await book.generateAsync({ type: "nodebuffer" }));

    const pkg = await Pkg.open(await zip.generateAsync({ type: "uint8array" }));
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "n");
    if (!prepared.ok) throw new Error(`the fixture was refused: ${prepared.why}`);
    const records = toRecordSet(ROWS);
    const plan = buildPlan(prepared.block, records, { runId: "n" });
    await runPlan(pkg, plan, records);
    return JSZip.loadAsync(await pkg.toBytes());
  }

  it("fills it, and both halves move together", async () => {
    const zip = await withValueCell('<c r="B2" t="inlineStr"><is><t>{{Revenue}}</t></is></c>');
    expect(await cachedValues(zip)).toEqual(["1250000"]);
    // The type goes with the text: a numeric cell still carrying `inlineStr`
    // is one Excel reads as a string, and the chart would plot nothing.
    expect((await sheetCells(zip))["B2"]).toEqual({ type: "n", value: "1250000" });
  });

  it("joins an inline string split across runs", async () => {
    // The same split a slide's paragraph gets, one format down: `<is>` holds
    // `<r><t>` runs whenever part of the cell is styled differently.
    const zip = await withValueCell('<c r="B2" t="inlineStr"><is><r><t>{{Reve</t></r><r><t>nue}}</t></r></is></c>');
    expect(await cachedValues(zip)).toEqual(["1250000"]);
    expect((await sheetCells(zip))["B2"]).toEqual({ type: "n", value: "1250000" });
  });

  it("counts a value the cache has no point for, instead of passing over it", async () => {
    /**
     * A cached point list is SPARSE by design: `<c:ptCount>` covers the range
     * and each `<c:pt>` carries an index, so a writer omits the point for a
     * cell it has no number for — which is exactly the cell somebody typed a
     * placeholder into. python-pptx, which authored this repo's own test-kit
     * chart, does that for a `None`; so does xlsxwriter.
     *
     * The numeric walk is over the POINTS, so such a cell was never opened:
     * not filled, not refused, not counted, not offered to the pane as a field.
     * The workbook's text pass fills it all the same, so the data sheet ends up
     * holding the row's figure under a chart with nowhere to draw it — and on a
     * host that refreshes the cache from the sheet, the bar appears later out
     * of nowhere.
     *
     * Counted rather than repaired, and the sentence carries the remedy: the
     * value is in the sheet, and Edit Data brings the chart into line. Writing
     * the missing point changes what PowerPoint draws, and whether the
     * documented Edit Data route produces a gap at all is unproven.
     */
    const deck = await makeDeck([
      { paragraphs: [["{{Name}}"]], chart: { categories: ["a", "b"], workbook: ["x"], values: ["10", "{{Revenue}}"] } },
    ]);
    const zip = await JSZip.loadAsync(deck);
    // Remove the cached point for index 1 — the placeholder's own cell — while
    // leaving `ptCount` at 2. That is the shape the writers above produce.
    const chart = await zip.file("ppt/charts/chart1.xml")!.async("string");
    // Scoped to the NUMBER cache. A chart holds a category cache too and it
    // comes first, so an unscoped replace takes a category point and the test
    // then measures nothing — which is what the first version of this did.
    const holed = chart.replace(/<c:numCache>[\s\S]*?<\/c:numCache>/, (cache) =>
      cache.replace(/<c:pt idx="1">[\s\S]*?<\/c:pt>/, ""),
    );
    expect(holed, "the fixture's cached points moved; this patches them by hand").not.toBe(chart);
    expect(holed, "the count still covers the cell").toContain('<c:ptCount val="2"/>');
    expect(holed, "and the placeholder's own point is the one gone").not.toContain("{{Revenue}}</c:v>");
    zip.file("ppt/charts/chart1.xml", holed);

    const pkg = await Pkg.open(await zip.generateAsync({ type: "uint8array" }));
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "h");
    if (!prepared.ok) throw new Error(prepared.why);
    const records = toRecordSet(ROWS);
    const out = await runPlan(pkg, buildPlan(prepared.block, records, { runId: "h" }), records);

    expect(out.graphics.numbers.filled, "there is no point to fill").toBe(0);
    expect(out.graphics.numbers.unplotted, "and the run says so").toBe(1);
    // NAMED as well as counted. The comment above says it was "not offered to
    // the pane as a field", and the next test is what that cost.
    expect(prepared.fields, "the pane is never offered the field").toContain("Revenue");
  });

  it("does not REFUSE a block whose only placeholder the cache has no point for", async () => {
    /**
     * `prepareBlock` refuses a block whose fields come back empty — "every copy
     * would be identical" — and the walk above never reported this one. So a
     * deck whose only placeholder is a chart value cell was refused outright,
     * with a sentence telling the author to go and type field names onto a
     * slide that already carried one. That is the documented workflow ("type
     * {{Revenue}} into a value cell the way you would type a number") meeting
     * the very cache shape the writers in the test above produce.
     *
     * The fixture is the same, minus the {{Name}} on the slide that was
     * carrying the block past the refusal.
     */
    const deck = await makeDeck([
      { paragraphs: [["Sales"]], chart: { categories: ["a", "b"], workbook: ["x"], values: ["10", "{{Revenue}}"] } },
    ]);
    const zip = await JSZip.loadAsync(deck);
    const chart = await zip.file("ppt/charts/chart1.xml")!.async("string");
    const holed = chart.replace(/<c:numCache>[\s\S]*?<\/c:numCache>/, (cache) =>
      cache.replace(/<c:pt idx="1">[\s\S]*?<\/c:pt>/, ""),
    );
    expect(holed, "the fixture's cached points moved; this patches them by hand").not.toBe(chart);
    zip.file("ppt/charts/chart1.xml", holed);

    const pkg = await Pkg.open(await zip.generateAsync({ type: "uint8array" }));
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "h");
    expect(prepared.ok, prepared.ok ? "" : prepared.why).toBe(true);
    if (!prepared.ok) return;
    expect(prepared.fields).toEqual(["Revenue"]);
  });

  it("holds an INLINE value cell it refused, which shares no string with anything", async () => {
    /**
     * The other half of holding a refused cell. A shared-string cell is held by
     * the `<si>` it reads through, because that is the node the workbook's text
     * pass reaches; a cell that carries its own text has no `<si>`, so it is
     * held by sheet and reference instead.
     *
     * Inline is what a generator writes when it never built a string table —
     * the same population `mergeGraphics` reads worksheets for at all — so this
     * is not a corner of the format, it is the other kind of producer.
     *
     * `{{Notes}}` against a `Notes` column holding a word: the numeric pass
     * refuses it, and the text pass must not then merge it.
     */
    const zip = await withValueCell('<c r="B2" t="inlineStr"><is><t>{{Notes}}</t></is></c>');
    const rels = await zip.file("ppt/charts/_rels/chart2.xml.rels")!.async("string");
    const target = /Target="([^"]*\.xlsx)"/.exec(rels)![1]!;
    const book = await JSZip.loadAsync(await zip.file(`ppt/${target.replace(/^\.\.\//, "")}`)!.async("nodebuffer"));
    const sheet = await book.file("xl/worksheets/sheet1.xml")!.async("string");
    expect(sheet, "the placeholder is what stays").toContain("{{Notes}}");
    expect(sheet, "and the row's words did not replace it").not.toContain("hello");
  });

  it("leaves a FORMULA cell and its formula alone", async () => {
    // `t="str"` is a formula's cached string result, and it keeps that result
    // in `<v>`. The reader used to name `str` alongside `inlineStr` and then
    // look for `<t>`, so it found nothing and skipped the cell — right
    // behaviour, false claim. Reading `<v>` instead would have made it merge,
    // and `writeNumber` clears every child of the cell: the `<f>` would go with
    // the placeholder. A merge may take a placeholder; it may not take a
    // formula.
    const zip = await withValueCell('<c r="B2" t="str"><f>A1</f><v>{{Revenue}}</v></c>');
    expect(await cachedValues(zip)).toEqual(["0"]);
    const rels = await zip.file("ppt/charts/_rels/chart2.xml.rels")!.async("string");
    const target = /Target="([^"]*\.xlsx)"/.exec(rels)![1]!;
    const book = await JSZip.loadAsync(await zip.file(`ppt/${target.replace(/^\.\.\//, "")}`)!.async("nodebuffer"));
    const sheet = await book.file("xl/worksheets/sheet1.xml")!.async("string");
    expect(sheet).toContain("<f>A1</f>");
    expect(sheet).toContain("{{Revenue}}");
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
    // The VALUE, which is what the promise is about. Asserting only the type
    // was vacuous against the defect this case describes: the workbook's text
    // pass runs after the numeric one and rewrites the shared string the
    // refused cell still points at — same type, different words. The manual
    // says `{{Notes}}` in a value cell "stays exactly as you typed it", and the
    // chart it belongs to still draws the template's number, so a workbook that
    // says something else is the two halves of one chart contradicting each
    // other — which is the half-merge `numbers.ts` exists to prevent.
    expect(cells["B2"]!.value, "the placeholder is what stays").toBe("{{Notes}}");
  });

  it("holds the whole shared string, and that costs a label that shares it", async () => {
    /**
     * The trade this makes, stated so it is a decision rather than a surprise.
     *
     * Excel keeps one `<si>` per distinct string, so a workbook where the same
     * placeholder appears in a value cell AND in a label cell has both pointing
     * at one entry. Holding it back keeps the value cell honest and leaves the
     * label unmerged.
     *
     * That is the right way round. A refused value cell whose sheet says
     * something else is a chart contradicting its own data — invisible until
     * somebody opens Edit Data, and it changes the drawing when they close it.
     * A label still reading `{{Notes}}` is wrong in a way the author can see on
     * the slide, which is the outcome this engine chooses everywhere else.
     */
    const { out, zip } = await merge({
      // Both cells hold the same text, so the fixture's workbook gives them one
      // shared entry — the case Excel produces by deduplicating.
      workbook: ["{{Notes}}"],
      values: ["{{Notes}}", "42"],
      categories: ["{{Name}}", "Everyone else"],
    });
    expect(out.graphics.numbers.refused).toBe(1);
    const cells = await sheetCells(zip);
    expect(cells["B2"]!.value, "the value cell keeps what the author typed").toBe("{{Notes}}");
  });

  it("and a value cell left EMPTY is not blanked in the sheet either", async () => {
    // The commoner shape, and the default policy. `onEmpty: "blank"` writes an
    // empty string for a missing cell, which through the workbook's text pass
    // emptied the shared string the refused value cell points at — so Edit Data
    // showed nothing where the placeholder had been, under a chart still
    // drawing the template's number.
    const { out, zip } = await merge(
      { workbook: ["{{Name}}"], values: ["{{Revenue}}", "42"], categories: ["{{Name}}", "Everyone else"] },
      [
        ["Name", "Revenue"],
        ["Ada", ""],
      ],
    );
    expect(out.graphics.numbers.refused).toBe(1);
    const cells = await sheetCells(zip);
    expect(cells["B2"]!.value, "the placeholder is what stays").toBe("{{Revenue}}");
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

  it("counts a series whose sheet the workbook does not have, rather than skipping it in silence", async () => {
    /**
     * Giving up is right — there is no safe guess between two sheets — but it
     * was a bare `continue`: no fill, no refusal, nothing. So a chart whose
     * values were never even looked at reported the same two zeros as a chart
     * with no values at all, and `summary.ts` reasoned from that pair that a
     * zero fill always means a refusal. It did not, and this was the case with
     * no signal of any kind.
     */
    const pkg = await Pkg.open(
      await makeDeck([
        {
          paragraphs: [["{{Name}}"]],
          chart: { categories: ["{{Name}}", "Other"], workbook: ["{{Name}}"], values: ["{{Revenue}}", "42"] },
        },
      ]),
    );
    // Two sheets declared and the formula naming neither, which is what stops
    // the single-sheet fallback from carrying it.
    const S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
    const embedding = pkg.partNames().find((n) => /^ppt\/embeddings\/.+\.xlsx$/.test(n))!;
    const book = await JSZip.loadAsync(await pkg.bytes(embedding));
    const rels = await book.file("xl/_rels/workbook.xml.rels")!.async("string");
    book.file(
      "xl/workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<workbook ${S} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>` +
        `<sheet name="Alpha" sheetId="1" r:id="${/Id="([^"]+)"[^>]*worksheet/.exec(rels)?.[1] ?? "rId1"}"/>` +
        `<sheet name="Beta" sheetId="2" r:id="${/Id="([^"]+)"[^>]*worksheet/.exec(rels)?.[1] ?? "rId1"}"/>` +
        `</sheets></workbook>`,
    );
    pkg.setBytes(embedding, await book.generateAsync({ type: "uint8array" }));

    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "s");
    if (!prepared.ok) throw new Error(prepared.why);
    const records = toRecordSet(ROWS);
    const out = await runPlan(pkg, buildPlan(prepared.block, records, { runId: "s" }), records);

    expect(out.graphics.numbers.filled, "nothing could be filled").toBe(0);
    expect(out.graphics.numbers.unreadable, "and the run says so").toBeGreaterThan(0);
  });

  it.each([
    ["exactly as declared", "Sheet2!$B$2:$B$3"],
    // Excel sheet names are case-INSENSITIVE — a workbook cannot hold both
    // `Data` and `data` — and the lookup was a plain `Map`, which is not. A
    // formula spelling the title differently from the declaration found
    // nothing, wrote nothing, and counted nothing. The single-sheet fallback
    // hides it completely, so it only appears on a workbook somebody added a
    // sheet to, which is why no round has met it.
    ["in a different case", "SHEET2!$B$2:$B$3"],
  ])("writes into the sheet the chart names %s", async (_what, formula) => {
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
      (await pkg.text("ppt/charts/chart1.xml")).replace("Sheet1!$B$2:$B$3", formula),
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
      unreadable: 0,
      unplotted: 0,
    });
  });

  it("fills them when the part is not under xl/worksheets at all", async () => {
    // The relationship target decides the path. Nothing in the format says a
    // worksheet lives in that folder either.
    const out = await mergeRenamed("xl/data.xml");
    expect(out.graphics.numbers).toEqual({ filled: 1, refused: 0, unreadable: 0, unplotted: 0 });
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
