/**
 * Charts and SmartArt, merged.
 *
 * Their text is DrawingML, so the READER was never the problem — `fieldsIn` has
 * reported these placeholders since it was written. What was missing is that
 * the parts holding them are shared by every clone, and that the same string is
 * kept in more than one place: a chart's labels in its cache and in the
 * workbook behind it, a SmartArt node's text in the model and in the rendering.
 *
 * Every assertion here is made against the FINISHED package — reopened from the
 * bytes, walked by relationship — rather than against the objects that wrote
 * it. A test that asks the writer what it wrote cannot see a relationship that
 * points at the wrong part.
 */
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { buildPlan } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import { toRecordSet } from "../src/core/data/recordset.js";
import { Pkg } from "../src/core/pptx/pkg.js";
import { REL_TYPE } from "../src/core/pptx/parts.js";
import { A_NS, C_NS, PKG_REL_NS, SSML_NS, elements, parseXml } from "../src/core/pptx/xml.js";
import { sheetNamed, withinInflatedBudget, workbookParts } from "../src/core/merge/workbook.js";
import { makeDeck, type SlideSpec } from "./fixtures/deck.js";

const ROWS = "Name\tRegion\nAda\tNordics\nGrace\tBenelux";

async function merge(spec: SlideSpec, rows = ROWS) {
  const pkg = await Pkg.open(await makeDeck([spec]));
  const records = toRecordSet(rows.split("\n").map((line) => line.split("\t")));
  const block = { id: "r1", slides: [{ path: "ppt/slides/slide1.xml", seq: 1, fields: [] }] };
  const plan = buildPlan(block, records, { runId: "r1" });
  const out = await runPlan(pkg, plan, records);
  // Reopened from the bytes the run produced. Everything below walks THIS zip.
  return { out, zip: await JSZip.loadAsync(await pkg.toBytes()) };
}

/** The parts one slide relates to, by relationship type suffix. */
async function related(zip: JSZip, part: string, endsWith: string): Promise<string[]> {
  const dir = part.slice(0, part.lastIndexOf("/"));
  const rels = await zip.file(`${dir}/_rels/${part.slice(part.lastIndexOf("/") + 1)}.rels`)?.async("string");
  if (!rels) return [];
  return elements(parseXml(rels), PKG_REL_NS, "Relationship")
    .filter((r) => (r.getAttribute("Type") ?? "").endsWith(endsWith))
    .map((r) => {
      const target = r.getAttribute("Target") ?? "";
      const segments = dir.split("/");
      for (const seg of target.split("/")) {
        if (seg === "..") segments.pop();
        else if (seg !== ".") segments.push(seg);
      }
      return segments.join("/");
    });
}

async function textOf(zip: JSZip, part: string, ns: string, local: string): Promise<string[]> {
  const xml = await zip.file(part)?.async("string");
  if (!xml) return [];
  return elements(parseXml(xml), ns, local).map((n) => n.textContent ?? "");
}

describe("a chart on a merged slide", () => {
  const withChart: SlideSpec = {
    paragraphs: [["Cover"]],
    chart: { title: "Revenue in {{Region}}", categories: ["{{Region}}", "Other"] },
  };

  it("fills the placeholder in the chart's title", async () => {
    const { zip } = await merge(withChart);
    const chart = (await related(zip, "ppt/slides/slide2.xml", "/chart"))[0] ?? "";
    expect((await textOf(zip, chart, A_NS, "t")).join("")).toBe("Revenue in Nordics");
  });

  it("fills the category labels, which are not paragraphs at all", async () => {
    // The labels a user actually writes are `<c:v>` inside a string cache. A
    // merge that only walked `<a:p>` would fill the title and leave every bar
    // on the chart named `{{Region}}`.
    const { zip } = await merge(withChart);
    const chart = (await related(zip, "ppt/slides/slide2.xml", "/chart"))[0] ?? "";
    const cache = elements(parseXml((await zip.file(chart)?.async("string")) ?? ""), C_NS, "strCache");
    expect(cache.flatMap((c) => elements(c, C_NS, "v").map((v) => v.textContent))).toEqual(["Nordics", "Other"]);
  });

  it("leaves a number cache alone, even when somebody has put a field in one", async () => {
    // `<c:v>` is the same element in both caches and only one of them holds
    // text. A merge that took every `<c:v>` would be writing into the values a
    // chart PLOTS, where the content has to parse as a number — so filling one
    // with "Nordics" produces a chart PowerPoint reads as corrupt data rather
    // than a chart with a merged bar.
    //
    // Written with a placeholder actually in the cache, because the ordinary
    // fixture's numbers contain none: a version of this test against "1" and
    // "2" passes whether or not the scoping is there, which is what it did.
    const { zip } = await merge({ ...withChart, chart: { categories: ["{{Region}}"], values: ["{{Region}}"] } });
    const chart = (await related(zip, "ppt/slides/slide2.xml", "/chart"))[0] ?? "";
    const doc = parseXml((await zip.file(chart)?.async("string")) ?? "");
    const nums = elements(doc, C_NS, "numCache").flatMap((c) => elements(c, C_NS, "v").map((v) => v.textContent));
    const strs = elements(doc, C_NS, "strCache").flatMap((c) => elements(c, C_NS, "v").map((v) => v.textContent));
    expect(nums).toEqual(["{{Region}}"]);
    // And the same field one element over IS filled, so this is scoping rather
    // than the merge having missed the chart altogether.
    expect(strs).toEqual(["Nordics"]);
  });

  it("gives every copy its own chart part", async () => {
    // The defect this whole pass exists for. A clone inherits the template's
    // relationships wholesale, so without cloning, both records merge into ONE
    // chart and the deck shows the last one's labels on every slide.
    const { zip } = await merge(withChart);
    const first = (await related(zip, "ppt/slides/slide2.xml", "/chart"))[0];
    const second = (await related(zip, "ppt/slides/slide3.xml", "/chart"))[0];
    expect(first).not.toBe(second);
    expect((await textOf(zip, first ?? "", A_NS, "t")).join("")).toBe("Revenue in Nordics");
    expect((await textOf(zip, second ?? "", A_NS, "t")).join("")).toBe("Revenue in Benelux");
  });

  it("leaves the template's own chart as the author wrote it", async () => {
    const { zip } = await merge(withChart);
    expect((await textOf(zip, "ppt/charts/chart1.xml", A_NS, "t")).join("")).toBe("Revenue in {{Region}}");
  });

  it("declares each copy's content type, or PowerPoint calls the file damaged", async () => {
    const { zip } = await merge(withChart);
    const types = (await zip.file("[Content_Types].xml")?.async("string")) ?? "";
    for (const slide of ["ppt/slides/slide2.xml", "ppt/slides/slide3.xml"]) {
      const chart = (await related(zip, slide, "/chart"))[0] ?? "";
      expect(types, `${chart} is not declared`).toContain(`PartName="/${chart}"`);
    }
  });
});

describe("the workbook behind a chart", () => {
  const withWorkbook: SlideSpec = {
    paragraphs: [["Cover"]],
    chart: { categories: ["{{Region}}"], workbook: ["{{Region}}", "Total"] },
  };

  /** The shared strings of the workbook this slide's chart reads. */
  async function strings(zip: JSZip, slide: string): Promise<string[]> {
    const chart = (await related(zip, slide, "/chart"))[0] ?? "";
    const path = (await related(zip, chart, "/package"))[0] ?? "";
    const bytes = await zip.file(path)?.async("uint8array");
    if (!bytes) return [];
    const book = await JSZip.loadAsync(bytes);
    const sst = (await book.file("xl/sharedStrings.xml")?.async("string")) ?? "";
    return elements(parseXml(sst), SSML_NS, "si").map((si) =>
      elements(si, SSML_NS, "t")
        .map((t) => t.textContent ?? "")
        .join(""),
    );
  }

  it("is merged too, or Edit Data reverts the chart in front of the user", async () => {
    // The cache is what PowerPoint draws; the workbook is what Excel opens on
    // "Edit Data", and closing that Excel refreshes the cache FROM the
    // workbook. Merge one and not the other and the deck is right until
    // somebody clicks the button.
    const { zip } = await merge(withWorkbook);
    expect(await strings(zip, "ppt/slides/slide2.xml")).toEqual(["Nordics", "Total"]);
  });

  it("finds a shared string Excel split into runs", async () => {
    // The fixture splits its first string the way an edited one is stored. A
    // per-node search finds nothing there, which is the same defect the slide
    // reader was written for.
    const { zip } = await merge({ ...withWorkbook, chart: { workbook: ["{{Region}}"] } });
    expect(await strings(zip, "ppt/slides/slide2.xml")).toEqual(["Nordics"]);
  });

  it("gives every copy its own workbook", async () => {
    const { zip } = await merge(withWorkbook);
    expect(await strings(zip, "ppt/slides/slide2.xml")).toEqual(["Nordics", "Total"]);
    expect(await strings(zip, "ppt/slides/slide3.xml")).toEqual(["Benelux", "Total"]);
  });

  it("stays a readable zip after being written back", async () => {
    // Written back through a second zip writer, inside a zip. A workbook that
    // comes out unopenable is a chart PowerPoint reports as corrupt data.
    const { zip } = await merge(withWorkbook);
    const chart = (await related(zip, "ppt/slides/slide2.xml", "/chart"))[0] ?? "";
    const path = (await related(zip, chart, "/package"))[0] ?? "";
    const book = await JSZip.loadAsync((await zip.file(path)?.async("uint8array")) as Uint8Array);
    expect(Object.keys(book.files)).toContain("xl/workbook.xml");
  });

  it("counts what it merged, and reports nothing unreadable", async () => {
    const { out } = await merge(withWorkbook);
    expect(out.graphics.workbooks).toBe(2);
    expect(out.graphics.unreadable).toEqual([]);
  });

  it("merges the chart anyway when the workbook cannot be opened", async () => {
    // An embedded part that is not a zip: another tool's OLE object under the
    // same relationship type. The chart's own cache is what the reader sees,
    // so losing 240 slides over it would be the wrong trade.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["Cover"]], chart: withWorkbook.chart }]));
    pkg.setBytes("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx", new Uint8Array([1, 2, 3, 4]));
    const records = toRecordSet(ROWS.split("\n").map((l) => l.split("\t")));
    const plan = buildPlan({ id: "r1", slides: [{ path: "ppt/slides/slide1.xml", seq: 1, fields: [] }] }, records, {
      runId: "r1",
    });
    const out = await runPlan(pkg, plan, records);

    expect(out.graphics.unreadable).toHaveLength(2);
    const zip = await JSZip.loadAsync(await pkg.toBytes());
    const chart = (await related(zip, "ppt/slides/slide2.xml", "/chart"))[0] ?? "";
    const cache = elements(parseXml((await zip.file(chart)?.async("string")) ?? ""), C_NS, "strCache");
    expect(cache.flatMap((c) => elements(c, C_NS, "v").map((v) => v.textContent))).toEqual(["Nordics"]);
  });
});

/**
 * A workbook written by a TOOL rather than by Excel.
 *
 * Two things such a file does differently, and the merge had a defect for each.
 * It puts a cell's string INLINE — `<c t="inlineStr"><is><t>` — instead of
 * pointing at a shared-string table it never built; and it names its parts
 * whatever it likes, because a part name in a package is arbitrary and only the
 * workbook's own declarations say where its sheets are.
 *
 * Both cells are patched into one sheet on purpose. A chart's VALUE cell is
 * read by the numeric pass and its LABEL cell by the text pass, so a file where
 * one is filled and the other is not is the two passes disagreeing about the
 * same workbook — which is what the shared reader in `workbook.ts` exists to
 * make impossible.
 */
describe("a workbook a generator wrote", () => {
  /** The fixture's workbook with both cells inline, and its sheet at `sheetPath`. */
  async function mergedSheet(sheetPath: string): Promise<string> {
    const deck = await makeDeck([
      { paragraphs: [["{{Name}}"]], chart: { categories: ["a"], workbook: ["x"], values: ["0"] } },
    ]);
    const zip = await JSZip.loadAsync(deck);
    const embedding = Object.keys(zip.files).find((n) => /embeddings\/.*\.xlsx$/.test(n)) ?? "";
    const book = await JSZip.loadAsync(await (zip.file(embedding) as JSZip.JSZipObject).async("nodebuffer"));

    const original = "xl/worksheets/sheet1.xml";
    let sheet = await (book.file(original) as JSZip.JSZipObject).async("string");
    const before = sheet;
    sheet = sheet.replace('<c r="B2"><v>0</v></c>', '<c r="B2" t="inlineStr"><is><t>{{Revenue}}</t></is></c>');
    sheet = sheet.replace(/<c r="A2"[^>]*>[\s\S]*?<\/c>/, '<c r="A2" t="inlineStr"><is><t>{{Name}}</t></is></c>');
    expect(sheet, "the fixture's cells moved; this test patches them by hand").not.toBe(before);

    if (sheetPath !== original) {
      book.remove(original);
      const rels = await (book.file("xl/_rels/workbook.xml.rels") as JSZip.JSZipObject).async("string");
      const moved = rels.replace('Target="worksheets/sheet1.xml"', `Target="${sheetPath.slice("xl/".length)}"`);
      expect(moved, "the fixture stopped naming its sheet by relationship").not.toBe(rels);
      book.file("xl/_rels/workbook.xml.rels", moved);
      const types = await (book.file("[Content_Types].xml") as JSZip.JSZipObject).async("string");
      book.file("[Content_Types].xml", types.replace(`/${original}`, `/${sheetPath}`));
    }
    book.file(sheetPath, sheet);
    zip.file(embedding, await book.generateAsync({ type: "nodebuffer" }));

    const pkg = await Pkg.open(await zip.generateAsync({ type: "uint8array" }));
    const records = toRecordSet([
      ["Name", "Revenue"],
      ["Ada", "1250000"],
    ]);
    const block = { id: "w", slides: [{ path: "ppt/slides/slide1.xml", seq: 1, fields: [] }] };
    await runPlan(pkg, buildPlan(block, records, { runId: "w" }), records);

    const out = await JSZip.loadAsync(await pkg.toBytes());
    const chart = (await related(out, "ppt/slides/slide2.xml", "/chart"))[0] ?? "";
    const workbook = (await related(out, chart, "/package"))[0] ?? "";
    const merged = await JSZip.loadAsync(await (out.file(workbook) as JSZip.JSZipObject).async("nodebuffer"));
    return (merged.file(sheetPath) as JSZip.JSZipObject).async("string");
  }

  it("fills a cell that holds its string inline", async () => {
    // `<is>` was not a text group, so `mergeWorkbook` opened every worksheet of
    // every embedded workbook and could never find anything in one. Its own
    // comment said the worksheets were read for exactly this case.
    const sheet = await mergedSheet("xl/worksheets/sheet1.xml");
    expect(sheet, "the value cell").not.toContain("{{Revenue}}");
    expect(sheet, "the label cell").not.toContain("{{Name}}");
    expect(sheet).toContain("Ada");
  });

  it("finds the sheet by what the workbook declares, not by its part name", async () => {
    // The numeric pass reads the declarations and the text pass matched
    // `xl/worksheets/sheetN.xml`, so this same file came back with its value
    // cell filled and the label cell beside it still reading `{{Name}}`.
    const sheet = await mergedSheet("xl/sheets/data.xml");
    expect(sheet, "the value cell").not.toContain("{{Revenue}}");
    expect(sheet, "the label cell").not.toContain("{{Name}}");
    expect(sheet).toContain("Ada");
  });
});

/**
 * A part whose NAME needs a percent escape.
 *
 * A relationship `Target` is a URI reference, so a part called `my chart.xml`
 * is written `../charts/my%20chart.xml`; a zip entry name is the literal name.
 * The two only meet if the resolver decodes, and it did not — so `pkg.has`
 * answered no for a part that is right there, and every pass that asks it
 * stepped over the chart in silence.
 *
 * `cloneSlideGraphics` is where that costs the deck. It skips what it cannot
 * find, so every merged copy keeps the TEMPLATE's chart relationship, and all
 * of them show the last record's labels — the shared-part defect this project
 * has now found for notes pages, comments, charts and diagrams.
 */
/**
 * Reading a workbook the way it describes itself.
 *
 * Driven directly rather than through a merge, because two of its rules cannot
 * be reached from a package this project can author: a workbook whose main part
 * is not `xl/workbook.xml`, and one that declares two sheets under one title.
 * A rule no test can drive is a rule nobody can check — and a mutation sweep
 * proved both were exactly that.
 */
describe("an embedded workbook that would inflate to gigabytes", () => {
  /**
   * A `.pptx` arrives from wherever the user got it, and a chart's data is a
   * whole `.xlsx` sitting inside it. Deflate reaches about 1000:1 on repetitive
   * XML — measured, 19 KB became 20 MB — and BOTH passes over a workbook open
   * it once per merged row, because every clone gets its own copy. At 240 rows
   * that is a task pane doing gigabytes of work it cannot be interrupted out
   * of.
   *
   * The zip declares each entry's inflated size, so the question is answered
   * before anything is inflated and costs nothing on an ordinary deck.
   *
   * Refusing is the answer an unparseable workbook already gets: the chart
   * keeps its cached values, the run finishes, and the pane says the data
   * behind it could not be opened — a sentence the user can act on, where a
   * frozen tab is not.
   */
  it("is refused rather than inflated", async () => {
    const book = new JSZip();
    book.file("xl/workbook.xml", `<workbook><sheets/></workbook>`);
    // Bigger than the budget, and it costs a few kilobytes on disk.
    book.file("xl/sharedStrings.xml", `<sst>${"a".repeat(80 * 1024 * 1024)}</sst>`);
    const bytes = await book.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    expect(bytes.length, "a bomb is small on disk").toBeLessThan(1024 * 1024);
    expect(withinInflatedBudget(await JSZip.loadAsync(bytes)), "refused").toBe(false);
  });

  it("and an ordinary one is not", async () => {
    // The other half, because a budget that refuses everything is not a budget.
    const book = new JSZip();
    book.file("xl/workbook.xml", `<workbook><sheets/></workbook>`);
    book.file("xl/sharedStrings.xml", `<sst>${"<si><t>Ada Lovelace</t></si>".repeat(20000)}</sst>`);
    const bytes = await book.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    expect(withinInflatedBudget(await JSZip.loadAsync(bytes)), "a real workbook passes").toBe(true);
  });
});

describe("workbookParts", () => {
  const REL = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
  const S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
  const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  /** A workbook whose main part is `main`, declaring the given sheets. */
  function book(main: string, sheets: [string, string][]): JSZip {
    const zip = new JSZip();
    zip.file(
      "_rels/.rels",
      `<Relationships ${REL}>` +
        `<Relationship Id="rIdX" Type="${TYPE}/extended-properties" Target="docProps/app.xml"/>` +
        `<Relationship Id="rIdM" Type="${TYPE}/officeDocument" Target="${main}"/>` +
        `</Relationships>`,
    );
    zip.file("docProps/app.xml", "<Properties/>");
    zip.file(
      main,
      `<workbook ${S} ${R}><sheets>` +
        sheets.map(([name], i) => `<sheet name="${name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("") +
        `</sheets></workbook>`,
    );
    const dir = main.slice(0, main.lastIndexOf("/"));
    zip.file(
      `${dir}/_rels/${main.slice(main.lastIndexOf("/") + 1)}.rels`,
      `<Relationships ${REL}>` +
        sheets
          .map(([, part], i) => `<Relationship Id="rId${i + 1}" Type="${TYPE}/worksheet" Target="${part}"/>`)
          .join("") +
        `<Relationship Id="rIdS" Type="${TYPE}/sharedStrings" Target="strings.xml"/>` +
        `</Relationships>`,
    );
    for (const [, part] of sheets) zip.file(`${dir}/${part}`, `<worksheet ${S}/>`);
    zip.file(`${dir}/strings.xml`, `<sst ${S}/>`);
    return zip;
  }

  it("follows the officeDocument relationship rather than a fixed name", async () => {
    const parts = await workbookParts(book("book/main.xml", [["Data", "sheets/one.xml"]]));
    expect(parts.sheets).toEqual(["book/sheets/one.xml"]);
    expect(sheetNamed(parts, "Data")).toBe("book/sheets/one.xml");
    expect(parts.sharedStrings).toBe("book/strings.xml");
  });

  it("keeps the FIRST sheet declared under a title, never the last", async () => {
    // Two sheets cannot share a title in a workbook Excel will open. If a
    // generator writes one anyway, a chart's `Sheet1!$B$2` has to reach the
    // sheet the workbook names first — taking the later one would fill a
    // different cell and the chart would plot the one that was left.
    const parts = await workbookParts(
      book("xl/workbook.xml", [
        ["Sheet1", "worksheets/first.xml"],
        ["Sheet1", "worksheets/second.xml"],
      ]),
    );
    expect(sheetNamed(parts, "Sheet1")).toBe("xl/worksheets/first.xml");
    // Both are still sheets — only the TITLE is claimed once.
    expect(parts.sheets).toEqual(["xl/worksheets/first.xml", "xl/worksheets/second.xml"]);
  });

  it("answers empty for an embedding that is not a workbook at all", async () => {
    // An OLE object under the same `package` relationship. Reported as nothing
    // to fill rather than thrown on: a merge does not lose 240 slides over one.
    const parts = await workbookParts(new JSZip());
    expect(parts.sheets).toEqual([]);
    expect(parts.sharedStrings).toBeUndefined();
  });
});

describe("a chart part whose name needs a percent escape", () => {
  it("still gets its own copy per merged slide", async () => {
    const deck = await makeDeck([{ paragraphs: [["{{Region}}"]], chart: { title: "{{Region}}" } }]);
    const zip = await JSZip.loadAsync(deck);
    const body = await (zip.file("ppt/charts/chart1.xml") as JSZip.JSZipObject).async("string");
    zip.remove("ppt/charts/chart1.xml");
    zip.file("ppt/charts/my chart.xml", body);
    const rels = await (zip.file("ppt/slides/_rels/slide1.xml.rels") as JSZip.JSZipObject).async("string");
    const escaped = rels.replace('Target="../charts/chart1.xml"', 'Target="../charts/my%20chart.xml"');
    expect(escaped, "the fixture stopped naming its chart the way this patches it").not.toBe(rels);
    zip.file("ppt/slides/_rels/slide1.xml.rels", escaped);
    const types = await (zip.file("[Content_Types].xml") as JSZip.JSZipObject).async("string");
    zip.file("[Content_Types].xml", types.replace("/ppt/charts/chart1.xml", "/ppt/charts/my%20chart.xml"));

    const pkg = await Pkg.open(await zip.generateAsync({ type: "uint8array" }));
    const records = toRecordSet(ROWS.split("\n").map((line) => line.split("\t")));
    const block = { id: "e", slides: [{ path: "ppt/slides/slide1.xml", seq: 1, fields: [] }] };
    const result = await runPlan(pkg, buildPlan(block, records, { runId: "e" }), records);
    const out = await JSZip.loadAsync(await pkg.toBytes());

    const charts = await Promise.all(result.slides.map((slide) => related(out, slide, "/chart")));
    expect(charts[0], "the copy kept the template's chart").not.toEqual(charts[1]);
    // And each one says its own row, which is what sharing a part destroys.
    const said = await Promise.all(charts.map(async (c) => (await textOf(out, c[0] ?? "", A_NS, "t")).join("")));
    expect(said).toEqual(["Nordics", "Benelux"]);
  });
});

describe("SmartArt the way PowerPoint actually writes it", () => {
  /**
   * The same diagram, with the `diagramDrawing` relationship on the SLIDE and
   * no `dataN.xml.rels` at all — which is what PowerPoint produced when a Basic
   * Process diagram was inserted by hand on 2026-08-28.
   *
   * Until then the engine looked for that relationship on the data part only,
   * and every fixture here was built to match. So a real diagram merged its
   * model, left its rendering untouched, and handed three copies one shared
   * drawing still reading `{{Region}}` — with the whole suite green, because
   * the reader and the fixtures shared one misreading.
   */
  const powerPointShape: SlideSpec = {
    paragraphs: [["Cover"]],
    smartArt: ["{{Name}} of {{Region}}", "Second"],
    smartArtDrawingOn: "slide",
  };

  /** The rendering, found where PowerPoint puts it: on the slide. */
  const drawingOf = async (zip: JSZip, slide: string) => (await related(zip, slide, "/diagramDrawing"))[0] ?? "";

  it("has no dataN.xml.rels, which is the condition that hid this", async () => {
    // Asserted so the fixture cannot quietly drift back to the shape that
    // agreed with the old reader.
    const { zip } = await merge(powerPointShape);
    expect(Object.keys(zip.files).filter((n) => /diagrams\/_rels\/data\d+\.xml\.rels/.test(n))).toEqual([]);
  });

  it("fills the drawing the slide points at", async () => {
    const { zip } = await merge(powerPointShape);
    const drawing = await drawingOf(zip, "ppt/slides/slide2.xml");
    expect(drawing, "the merged slide names no drawing").not.toBe("");
    expect((await textOf(zip, drawing, A_NS, "t")).join("")).toContain("Ada of Nordics");
  });

  it("gives every copy its own drawing, with its own row in it", async () => {
    const { zip } = await merge(powerPointShape);
    const first = await drawingOf(zip, "ppt/slides/slide2.xml");
    const second = await drawingOf(zip, "ppt/slides/slide3.xml");
    expect(first).not.toBe(second);
    expect((await textOf(zip, first, A_NS, "t")).join("")).toContain("Ada of Nordics");
    expect((await textOf(zip, second, A_NS, "t")).join("")).toContain("Grace of Benelux");
  });

  it("leaves no merged copy pointing at the template's own drawing", async () => {
    // The failure exactly as it appeared in the round: the copies were made,
    // the models were filled, and every one of them still named drawing1.
    const { zip } = await merge(powerPointShape);
    for (const slide of ["ppt/slides/slide2.xml", "ppt/slides/slide3.xml"]) {
      expect(await drawingOf(zip, slide), slide).not.toBe("ppt/diagrams/drawing1.xml");
    }
  });
});

describe("SmartArt on a merged slide", () => {
  const withSmartArt: SlideSpec = { paragraphs: [["Cover"]], smartArt: ["{{Name}} of {{Region}}", "Second"] };

  /** The model and the rendering, in that order, for one merged slide. */
  async function halves(zip: JSZip, slide: string): Promise<{ data: string; drawing: string }> {
    const data = (await related(zip, slide, "/diagramData"))[0] ?? "";
    const drawing = (await related(zip, data, "/diagramDrawing"))[0] ?? "";
    return { data, drawing };
  }

  it("fills the model", async () => {
    const { zip } = await merge(withSmartArt);
    const { data } = await halves(zip, "ppt/slides/slide2.xml");
    expect((await textOf(zip, data, A_NS, "t")).join("")).toContain("Ada of Nordics");
  });

  it("fills the DRAWING, which is the half anybody sees", async () => {
    // PowerPoint displays the laid-out rendering rather than re-running the
    // layout engine on open. A merge that filled only the model would produce a
    // deck whose SmartArt still reads `{{Name}}` on screen.
    const { zip } = await merge(withSmartArt);
    const { drawing } = await halves(zip, "ppt/slides/slide2.xml");
    expect((await textOf(zip, drawing, A_NS, "t")).join("")).toContain("Ada of Nordics");
  });

  it("gives every copy its own model and its own drawing", async () => {
    const { zip } = await merge(withSmartArt);
    const first = await halves(zip, "ppt/slides/slide2.xml");
    const second = await halves(zip, "ppt/slides/slide3.xml");
    expect(first.data).not.toBe(second.data);
    expect(first.drawing).not.toBe(second.drawing);
    expect((await textOf(zip, second.drawing, A_NS, "t")).join("")).toContain("Grace of Benelux");
  });

  it("keeps the layout, quick style and colours shared", async () => {
    // The other half of the rule. These are read-only styling, and copying them
    // per record would multiply a template's styling by the row count for no
    // change in what anybody sees.
    const { zip } = await merge(withSmartArt);
    for (const [type, part] of [
      ["/diagramLayout", "ppt/diagrams/layout1.xml"],
      ["/diagramQuickStyle", "ppt/diagrams/quickStyle1.xml"],
      ["/diagramColors", "ppt/diagrams/colors1.xml"],
    ] as const) {
      const first = (await related(zip, "ppt/slides/slide2.xml", type))[0];
      const second = (await related(zip, "ppt/slides/slide3.xml", type))[0];
      expect(first, type).toBe(second);
      // Named rather than merely equal: two clones agreeing on a part that is
      // itself a copy would satisfy the line above and still be per-record
      // styling.
      expect(first, type).toBe(part);
    }
  });

  it("declares the copies' content types", async () => {
    const { zip } = await merge(withSmartArt);
    const types = (await zip.file("[Content_Types].xml")?.async("string")) ?? "";
    const { data, drawing } = await halves(zip, "ppt/slides/slide3.xml");
    expect(types).toContain(`PartName="/${data}"`);
    expect(types).toContain(`PartName="/${drawing}"`);
  });
});

describe("what the run leaves behind", () => {
  /**
   * The template slides are taken out of the package before it is handed to
   * PowerPoint — inserting them would put the author's own placeholder slides
   * back into their deck on every run. Their charts and SmartArt have to go
   * with them, and only with them.
   */
  it("takes a removed template slide's chart out of the package too", async () => {
    const pkg = await Pkg.open(
      await makeDeck([{ paragraphs: [["Cover"]], chart: { title: "{{Region}}", workbook: ["{{Region}}"] } }]),
    );
    const records = toRecordSet(ROWS.split("\n").map((l) => l.split("\t")));
    const plan = buildPlan({ id: "r1", slides: [{ path: "ppt/slides/slide1.xml", seq: 1, fields: [] }] }, records, {
      runId: "r1",
    });
    await runPlan(pkg, plan, records);
    await pkg.removeSlide("ppt/slides/slide1.xml");

    const zip = await JSZip.loadAsync(await pkg.toBytes());
    expect(Object.keys(zip.files)).not.toContain("ppt/charts/chart1.xml");
    expect(Object.keys(zip.files)).not.toContain("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx");
    // And the copies are untouched: they are what the deck is for.
    const chart = (await related(zip, "ppt/slides/slide2.xml", "/chart"))[0] ?? "";
    expect(zip.file(chart)).not.toBeNull();
    expect(
      (await related(zip, chart, "/package"))[0] && zip.file((await related(zip, chart, "/package"))[0] ?? ""),
    ).not.toBeNull();
  });

  it("takes a modern chart's fallback picture out with the template slide", async () => {
    // The picture a modern chart carries for hosts too old to draw it. Every
    // merged copy replaces it with a notice and stops relating to it, so once
    // the template slide goes nothing points at the bytes — a rendering of the
    // TEMPLATE's figures, riding along in the file that gets sent out.
    //
    // Media used not to be a candidate for the sweep at all, so this survived
    // every merge. A picture is likelier than a chart to be shared, and the
    // referrer scan rather than the candidate list is what makes that safe —
    // which the next test is about.
    const pkg = await Pkg.open(
      await makeDeck([
        { paragraphs: [["Cover"]], modernChart: { title: "{{Region}}", categories: ["{{Region}}"] } },
        { paragraphs: [["after"]] },
      ]),
    );
    const records = toRecordSet(ROWS.split("\n").map((l) => l.split("\t")));
    const plan = buildPlan({ id: "r1", slides: [{ path: "ppt/slides/slide1.xml", seq: 1, fields: [] }] }, records, {
      runId: "r1",
    });
    await runPlan(pkg, plan, records);
    expect(pkg.partNames(), "the fixture drew no fallback picture, so this proves nothing").toContain(
      "ppt/media/chart1.png",
    );

    await pkg.removeSlide("ppt/slides/slide1.xml");
    expect(pkg.partNames()).not.toContain("ppt/media/chart1.png");
  });

  it("keeps a picture another part still points at", async () => {
    // The half that must NOT be swept, and the reason media can be a rule about
    // media in general. A logo on two slides, a photo used twice: a part any
    // other part still names is left exactly where it is, and a sweep that took
    // it would leave a surviving slide showing a missing picture.
    const pkg = await Pkg.open(
      await makeDeck([
        { paragraphs: [["Cover"]], modernChart: { title: "{{Region}}", categories: ["{{Region}}"] } },
        { paragraphs: [["after"]] },
      ]),
    );
    // A second referrer, exactly as a deck that used one image twice would have.
    await pkg.addRel("ppt/slides/slide2.xml", REL_TYPE.image, "../media/chart1.png");

    const records = toRecordSet(ROWS.split("\n").map((l) => l.split("\t")));
    const plan = buildPlan({ id: "r1", slides: [{ path: "ppt/slides/slide1.xml", seq: 1, fields: [] }] }, records, {
      runId: "r1",
    });
    await runPlan(pkg, plan, records);
    await pkg.removeSlide("ppt/slides/slide1.xml");

    expect(pkg.partNames(), "swept a picture a surviving slide still shows").toContain("ppt/media/chart1.png");
  });

  it("keeps a diagram's shared styling, which the copies still point at", async () => {
    // The half that must NOT be swept. Layout, quick style and colours are
    // shared with every clone by design, so removing them with the template
    // would leave every merged slide referencing a part that is not there —
    // which is how a sweep turns into a damaged file.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["Cover"]], smartArt: ["{{Region}}"] }]));
    const records = toRecordSet(ROWS.split("\n").map((l) => l.split("\t")));
    const plan = buildPlan({ id: "r1", slides: [{ path: "ppt/slides/slide1.xml", seq: 1, fields: [] }] }, records, {
      runId: "r1",
    });
    await runPlan(pkg, plan, records);
    await pkg.removeSlide("ppt/slides/slide1.xml");

    const zip = await JSZip.loadAsync(await pkg.toBytes());
    expect(Object.keys(zip.files)).not.toContain("ppt/diagrams/data1.xml");
    expect(Object.keys(zip.files)).not.toContain("ppt/diagrams/drawing1.xml");
    for (const shared of ["layout1", "quickStyle1", "colors1"]) {
      expect(zip.file(`ppt/diagrams/${shared}.xml`), shared).not.toBeNull();
    }
  });
});

describe("no two merged slides share a part the merge writes into", () => {
  /**
   * A sweep rather than a list, because the failure this catches is a part type
   * nobody remembered. Every per-record part is cloned by a named branch of
   * `cloneSlideGraphics`; a template carrying a kind that has no branch merges
   * every record into ONE part and the whole deck shows the last row — the
   * defect this file's header calls "shared by every clone", found three times
   * already in three different part types.
   *
   * Reachability rather than the branch list, so a part reached by a hop nobody
   * thought about is still covered.
   *
   * Layouts, masters, themes and a SmartArt's `layout`, `colors` and
   * `quickStyle` are deliberately NOT here. They are static definitions that no
   * pass writes into, and sharing them is as right as sharing the theme. Media
   * is shared on purpose too — one logo on 240 rows is one part.
   */
  const OWNED = /^ppt\/(charts|embeddings|notesSlides|tags)\/|^ppt\/diagrams\/(data|drawing)/;

  /** Every part reachable from a slide, however many hops out. */
  async function reach(pkg: Pkg, from: string): Promise<string[]> {
    const seen = new Set<string>();
    const queue = [from];
    while (queue.length) {
      for (const next of await pkg.relatedParts(queue.shift() as string)) {
        if (seen.has(next)) continue;
        seen.add(next);
        queue.push(next);
      }
    }
    return [...seen].filter((p) => OWNED.test(p));
  }

  it.each([["data"], ["slide"]] as const)("with the diagram drawing hung on the %s part", async (drawingOn) => {
    const pkg = await Pkg.open(
      await makeDeck([
        {
          paragraphs: [["{{Name}}"]],
          notes: "Call {{Name}} afterwards",
          chart: { title: "{{Name}}", categories: ["{{Name}}", "b"], workbook: ["{{Name}}"], values: ["1", "42"] },
          smartArt: ["{{Name}}", "second"],
          smartArtDrawingOn: drawingOn,
        },
        { paragraphs: [["after"]] },
      ]),
    );
    const records = toRecordSet([["Name"], ["Ada"], ["Bo"], ["Cy"]]);
    const block = { id: "r1", slides: [{ path: "ppt/slides/slide1.xml", seq: 1, fields: [] }] };
    const out = await runPlan(pkg, buildPlan(block, records, { runId: "r1" }), records);
    expect(out.slides).toHaveLength(3);

    const owned = new Map<string, string[]>();
    for (const slide of out.slides) owned.set(slide, await reach(pkg, slide));
    // Every slide must have brought something, or "nothing is shared" is only
    // a statement about an empty set.
    for (const [slide, parts] of owned) expect(parts.length, `${slide} reached no per-record part`).toBeGreaterThan(4);

    const shared: string[] = [];
    const slides = [...owned.keys()];
    for (let i = 0; i < slides.length; i++) {
      for (let j = i + 1; j < slides.length; j++) {
        const a = owned.get(slides[i] as string) as string[];
        const b = owned.get(slides[j] as string) as string[];
        for (const part of a) if (b.includes(part)) shared.push(`${slides[i]} and ${slides[j]} share ${part}`);
      }
    }
    expect(shared, "two records merge into one part, so the deck shows one of them twice").toEqual([]);
  });
});
