/**
 * Modern charts are NOT merged, and this pins that.
 *
 * A waterfall, funnel, treemap, sunburst, histogram or box-and-whisker chart is
 * not a `<c:chartSpace>`. PowerPoint writes it as a separate part reached
 * through a different relationship type — `…/2014/relationships/chartEx` — and
 * nothing in this engine knows that type. So such a chart is neither cloned per
 * copy nor filled: every merged slide points at the SAME part, still holding
 * the template's placeholders.
 *
 * That is the failure `docs/TEST-KIT.md` check 2 exists to catch, live for a
 * chart family this project does not handle. It is written down in the manual
 * and on the backlog rather than fixed here, because where a chartEx keeps its
 * labels is a fact about a schema nobody here has a real file of — and guessing
 * a schema is exactly what produced the SmartArt drawing bug, where the fixture
 * and the reader agreed with each other and disagreed with PowerPoint.
 *
 * When this test starts failing, somebody has added support. Read it then: it
 * says what the limit WAS.
 */
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { Pkg } from "../src/core/pptx/pkg.js";
import { prepareBlock } from "../src/core/merge/prepare.js";
import { buildPlan } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import { toRecordSet } from "../src/core/data/recordset.js";
import { makeDeck } from "./fixtures/deck.js";

const CHARTEX_REL = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
const CHARTEX_TYPE = "application/vnd.ms-office.chartex+xml";
const CX = 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

async function mergeWithModernChart() {
  const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["{{Name}}"]] }]));
  pkg.setText(
    "ppt/charts/chartEx1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<cx:chartSpace ${CX} ${A}>` +
      `<cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:lvl ptCount="2">` +
      `<cx:pt idx="0">{{Region}}</cx:pt><cx:pt idx="1">Everyone else</cx:pt>` +
      `</cx:lvl></cx:strDim></cx:data></cx:chartData>` +
      `<cx:chart><cx:title><cx:tx><cx:rich><a:bodyPr/><a:p><a:r><a:t>Revenue for {{Name}}</a:t></a:r></a:p>` +
      `</cx:rich></cx:tx></cx:title></cx:chart></cx:chartSpace>`,
  );
  await pkg.addContentTypeOverride("/ppt/charts/chartEx1.xml", CHARTEX_TYPE);
  await pkg.addRel("ppt/slides/slide1.xml", CHARTEX_REL, "../charts/chartEx1.xml");

  const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "cx");
  if (!prepared.ok) throw new Error(prepared.why);
  const records = toRecordSet([
    ["Name", "Region"],
    ["Ada", "Nordics"],
    ["Grace", "Benelux"],
  ]);
  await runPlan(pkg, buildPlan(prepared.block, records, { runId: "cx" }), records);
  return JSZip.loadAsync(await pkg.toBytes());
}

describe("a chart PowerPoint stores as chartEx", () => {
  it("is not given a copy of its own — every merged slide shares one", async () => {
    const zip = await mergeWithModernChart();
    const parts = Object.keys(zip.files).filter((n) => /^ppt\/charts\/chartEx\d+\.xml$/.test(n));
    expect(parts, "a copy per slide would mean this is supported now").toHaveLength(1);

    const targets: string[] = [];
    for (const s of Object.keys(zip.files).filter((n) => /slides\/slide\d+\.xml$/.test(n))) {
      const rels = (await zip.file(s.replace("slides/", "slides/_rels/") + ".rels")?.async("string")) ?? "";
      for (const m of rels.matchAll(/Target="([^"]*chartEx[^"]*)"/g)) targets.push(m[1]!);
    }
    expect(targets.length, "the template slide and both merged copies").toBe(3);
    expect(new Set(targets).size, "all three point at the same part").toBe(1);
  });

  it("keeps its placeholders, in both places it holds text", async () => {
    const zip = await mergeWithModernChart();
    const xml = await zip.file("ppt/charts/chartEx1.xml")!.async("string");
    // The category label, in `<cx:pt>` — a text node no reader here knows.
    expect(xml).toContain("{{Region}}");
    // And the title, which IS ordinary DrawingML: it would merge the moment the
    // part were visited at all. Nothing visits it, so the reader is not the
    // missing piece — reaching the part is.
    expect(xml).toContain("Revenue for {{Name}}");
  });
});
