/**
 * Modern charts — the chartEx family — merged.
 *
 * A waterfall, funnel, treemap, sunburst, histogram, pareto, box-and-whisker or
 * region map is not a `<c:chartSpace>`. PowerPoint writes it as a separate part
 * under `…/2014/relationships/chartEx`, and this file used to PIN the fact that
 * nothing here knew that type. Its header said to read it when it went red,
 * because red would mean somebody had added support. This is that.
 *
 * The limit it recorded, for the record: a merged deck carried ONE chartEx part
 * for the template and every copy, `prepareBlock` did not report the fields
 * inside it at all — so a block whose only placeholder was in the chart was
 * refused as empty — and every placeholder shipped.
 *
 * The fixture is shaped from a REAL PowerPoint file, `funnel-pp1.pptx` in
 * LibreOffice's chart test data, read on 2026-08-29. That matters more here
 * than anywhere else in this suite: several details of this format would have
 * been guessed wrong, and each is asserted below because each was observed.
 */
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { Pkg } from "../src/core/pptx/pkg.js";
import { prepareBlock } from "../src/core/merge/prepare.js";
import { buildPlan } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import { toRecordSet } from "../src/core/data/recordset.js";
import { A_NS, CX_NS, MC_NS, PKG_REL_NS, SSML_NS, child, children, elements, parseXml } from "../src/core/pptx/xml.js";
import { makeDeck, type ModernChartSpec, type SlideSpec } from "./fixtures/deck.js";

const ROWS = [
  ["Name", "Region", "Revenue"],
  ["Ada", "Nordics", "1250000"],
  ["Grace", "Benelux", "880000"],
];

const FUNNEL: ModernChartSpec = {
  title: "Pipeline for {{Name}}",
  categories: ["{{Region}}", "Everyone else"],
  series: "{{Name}}",
  workbook: ["{{Region}}", "Everyone else", "{{Name}}"],
};

const MERGED_SLIDES = ["ppt/slides/slide3.xml", "ppt/slides/slide4.xml"];

async function mergeDeck(spec: SlideSpec) {
  const pkg = await Pkg.open(await makeDeck([spec, { paragraphs: [["after"]] }]));
  const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "cx");
  if (!prepared.ok) throw new Error(`refused: ${prepared.why}`);
  const records = toRecordSet(ROWS);
  const result = await runPlan(pkg, buildPlan(prepared.block, records, { runId: "cx" }), records);
  return { prepared, result, zip: await JSZip.loadAsync(await pkg.toBytes()) };
}

/** The parts a slide relates to, by relationship type suffix. */
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

async function docOf(zip: JSZip, part: string): Promise<Document> {
  return parseXml((await zip.file(part)?.async("string")) ?? "");
}

/** The chartEx part one merged slide points at. */
async function chartOf(zip: JSZip, slide: string): Promise<Document> {
  return docOf(zip, (await related(zip, slide, "/chartEx"))[0] ?? "");
}

/** `<cx:pt>` under ONE kind of dim, which is the whole distinction that matters. */
function pointsIn(root: Document | Element, dim: "strDim" | "numDim"): string[] {
  return elements(root, CX_NS, dim).flatMap((d) => elements(d, CX_NS, "pt").map((p) => p.textContent ?? ""));
}

function textIn(root: Document | Element): string {
  return elements(root, A_NS, "t")
    .map((t) => t.textContent ?? "")
    .join("");
}

describe("a modern chart is merged", () => {
  it("reports the fields inside it, so the pane can offer a column for them", async () => {
    // The half that was worse than "not filled": these were invisible. The step
    // listing what is on the slides offered no Region column, and a block whose
    // only placeholder was in the chart was refused as having none.
    const { prepared } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: FUNNEL });
    expect(prepared.ok && [...prepared.fields].sort()).toEqual(["Name", "Region"]);
  });

  it("gives every copy its own part", async () => {
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: FUNNEL });
    const parts = Object.keys(zip.files).filter((n) => /^ppt\/charts\/chartEx\d+\.xml$/.test(n));
    expect(parts, "the template's, and one per record").toHaveLength(3);
    const first = (await related(zip, MERGED_SLIDES[0] ?? "", "/chartEx"))[0];
    const second = (await related(zip, MERGED_SLIDES[1] ?? "", "/chartEx"))[0];
    expect(first).not.toBe(second);
  });

  it("fills the category labels, which are cx:pt and not paragraphs", async () => {
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: FUNNEL });
    expect(pointsIn(await chartOf(zip, MERGED_SLIDES[0] ?? ""), "strDim")).toEqual(["Nordics", "Everyone else"]);
    expect(pointsIn(await chartOf(zip, MERGED_SLIDES[1] ?? ""), "strDim")).toEqual(["Benelux", "Everyone else"]);
  });

  it("leaves the plotted values alone, though they are the SAME element", async () => {
    // `<cx:pt>` in a `<cx:numDim>` is a number the chart plots. Filling one with
    // "Nordics" makes a chart PowerPoint reads as corrupt data, so this is
    // scoped by the dim exactly as `<c:v>` is scoped by str- versus numCache.
    // Written with a placeholder actually among the values: a version against
    // ordinary numbers passes whether or not the scoping is there.
    const spec = { ...FUNNEL, categories: ["{{Region}}"], values: ["{{Region}}"] };
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: spec });
    const doc = await chartOf(zip, MERGED_SLIDES[0] ?? "");
    expect(pointsIn(doc, "numDim")).toEqual(["{{Region}}"]);
    // And the same field one element over IS filled, so this is scoping rather
    // than the merge having missed the part altogether.
    expect(pointsIn(doc, "strDim")).toEqual(["Nordics"]);
  });

  it("fills a plotted VALUE from the workbook cell it was typed into", async () => {
    // The placeholder cannot go in `<cx:pt>` — that has to parse as a number,
    // and a chart carrying `{{Revenue}}` there is one PowerPoint reads as
    // corrupt. It goes where the value actually lives: the cell somebody typed
    // it into through Edit Data. Then BOTH copies of the number have to move —
    // the cell, which Excel shows, and the `<cx:lvl>` cache, which PowerPoint
    // draws from without opening the workbook at all.
    //
    // `<cx:f>` is what joins them, exactly as `<c:f>` does for a classic chart.
    const spec = { ...FUNNEL, values: ["{{Revenue}}", "42"] };
    const { result, zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: spec });
    expect(result.graphics.numbers).toEqual({ filled: 2, refused: 0, unreadable: 0, unplotted: 0 });
    expect(pointsIn(await chartOf(zip, MERGED_SLIDES[0] ?? ""), "numDim")).toEqual(["1250000", "42"]);
    expect(pointsIn(await chartOf(zip, MERGED_SLIDES[1] ?? ""), "numDim")).toEqual(["880000", "42"]);
  });

  it("turns that workbook cell back into a number, so Edit Data does not undo it", async () => {
    // A cell holding text plots nothing. Closing Excel refreshes the chart from
    // the workbook, so a merge that filled only the cache reverts the moment
    // the user clicks Edit Data and closes it again — the half-merge this
    // engine already knows about from the text side.
    const spec = { ...FUNNEL, values: ["{{Revenue}}", "42"] };
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: spec });
    const chart = (await related(zip, MERGED_SLIDES[0] ?? "", "/chartEx"))[0] ?? "";
    const book = (await related(zip, chart, "/package"))[0] ?? "";
    const inner = await JSZip.loadAsync((await zip.file(book)?.async("uint8array")) ?? new Uint8Array());
    const sheet = parseXml((await inner.file("xl/worksheets/sheet1.xml")?.async("string")) ?? "");
    const cell = elements(sheet, SSML_NS, "c").find((c) => c.getAttribute("r") === "B2");
    // No `t` at all is what makes a cell numeric, and the `<v>` is the number.
    expect(cell?.getAttribute("t")).toBeNull();
    expect(child(cell as Element, SSML_NS, "v")?.textContent).toBe("1250000");
  });

  it("reports a field that is only in a value cell, so the pane offers a column", async () => {
    // Reported from the same walk that fills them, not a second reader: a
    // scanner with its own opinion of which cells hold a placeholder is free to
    // disagree with the merge, and the pane would then offer a column that
    // fills nothing, or none for a cell that does.
    const spec = {
      ...FUNNEL,
      title: undefined,
      categories: ["one", "two"],
      series: undefined,
      values: ["{{Revenue}}", "42"],
    };
    const { prepared } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: spec });
    expect(prepared.ok && [...prepared.fields].sort()).toEqual(["Revenue"]);
  });

  it("refuses a value that will not be a number, and says so rather than guessing", async () => {
    // `{{Region}}` resolves to "Nordics", which is not a bar height. Guessing
    // zero would draw a chart the data never said, so the cache keeps what it
    // had and the run reports the refusal.
    const spec = { ...FUNNEL, values: ["{{Region}}", "42"] };
    const { result, zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: spec });
    expect(result.graphics.numbers).toEqual({ filled: 0, refused: 2, unreadable: 0, unplotted: 0 });
    expect(pointsIn(await chartOf(zip, MERGED_SLIDES[0] ?? ""), "numDim")).toEqual(["{{Region}}", "42"]);
  });

  it("fills the series name, which is cx:tx/cx:txData/cx:v", async () => {
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: FUNNEL });
    const doc = await chartOf(zip, MERGED_SLIDES[1] ?? "");
    const values = elements(doc, CX_NS, "txData").map((d) => child(d, CX_NS, "v")?.textContent ?? "");
    expect(values).toEqual(["Grace"]);
  });

  it("fills the title, which is DrawingML with no cx:tx beside it", async () => {
    // The trap that made reading a real file worth it. One chartEx keeps its
    // title as `<cx:v>` AND as a DrawingML copy; another keeps it only as
    // DrawingML. Both are filled, so neither shape merges by luck.
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: FUNNEL });
    expect(textIn(await chartOf(zip, MERGED_SLIDES[0] ?? ""))).toBe("Pipeline for Ada");
  });

  it("fills the workbook behind it, which Edit Data opens", async () => {
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: FUNNEL });
    const part = (await related(zip, MERGED_SLIDES[0] ?? "", "/chartEx"))[0] ?? "";
    const book = (await related(zip, part, "/package"))[0] ?? "";
    const inner = await JSZip.loadAsync((await zip.file(book)?.async("uint8array")) as Uint8Array);
    const sst = parseXml((await inner.file("xl/sharedStrings.xml")?.async("string")) ?? "");
    const strings = elements(sst, SSML_NS, "si").map((si) =>
      elements(si, SSML_NS, "t")
        .map((t) => t.textContent ?? "")
        .join(""),
    );
    expect(strings).toEqual(["Nordics", "Everyone else", "Ada"]);
  });

  it("leaves the template's own chart as the author wrote it", async () => {
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: FUNNEL });
    expect(pointsIn(await docOf(zip, "ppt/charts/chartEx1.xml"), "strDim")).toContain("{{Region}}");
  });

  it("declares each copy's content type, or PowerPoint calls the file damaged", async () => {
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: FUNNEL });
    const types = (await zip.file("[Content_Types].xml")?.async("string")) ?? "";
    for (const slide of MERGED_SLIDES) {
      const part = (await related(zip, slide, "/chartEx"))[0] ?? "";
      expect(types, `${part} is not declared`).toContain(`PartName="/${part}"`);
    }
  });
});

describe("the fallback branch a merged copy carries", () => {
  /** The `mc:Fallback` of the first AlternateContent on a slide. */
  async function fallbackOf(zip: JSZip, slide: string): Promise<Element | undefined> {
    const alternate = elements(await docOf(zip, slide), MC_NS, "AlternateContent")[0];
    return alternate ? child(alternate, MC_NS, "Fallback") : undefined;
  }

  it("no longer shows a picture of the template's data", async () => {
    // The decision, taken by the owner on 2026-08-29. PowerPoint writes a
    // RENDERED PNG here and regenerates it on save; this engine has no renderer
    // and cannot. Keeping the template's picture would put another recipient's
    // figures under this recipient's name, in the file that gets sent out.
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: FUNNEL });
    for (const slide of MERGED_SLIDES) {
      const fallback = await fallbackOf(zip, slide);
      expect(fallback, slide).toBeDefined();
      expect(elements(fallback as Element, A_NS, "blip"), `${slide} kept a picture`).toHaveLength(0);
    }
  });

  it("says why, rather than leaving a hole", async () => {
    // Dropping the branch is legal and silent. The point of writing a shape at
    // all is that somebody on an old host learns why the chart is missing.
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: FUNNEL });
    const said = textIn((await fallbackOf(zip, MERGED_SLIDES[0] ?? "")) as Element);
    expect(said).toContain("newer version of PowerPoint");
    expect(said.length).toBeGreaterThan(20);
  });

  it("centres the notice in the frame it inherited", async () => {
    // The shape is the size of the CHART, so an un-anchored sentence sits along
    // the top edge with a chart's worth of empty box under it — which reads as
    // something that failed to load rather than something saying why. Seen on
    // 2026-08-30 by forcing the branch with `test-kit/driver/force-fallback.mjs`,
    // because a host that takes it unaided is hard to come by.
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: FUNNEL });
    const fallback = (await fallbackOf(zip, MERGED_SLIDES[0] ?? "")) as Element;

    const bodyPr = elements(fallback, A_NS, "bodyPr");
    expect(bodyPr, "no text body to anchor").toHaveLength(1);
    expect(bodyPr[0]?.getAttribute("anchor"), "the notice is not centred vertically").toBe("ctr");

    const centred = elements(fallback, A_NS, "pPr").some((n) => n.getAttribute("algn") === "ctr");
    expect(centred, "the notice is not centred horizontally").toBe(true);
  });

  it("locks the notice against editing", async () => {
    // Somebody on a host too old to draw the chart must not be able to type
    // into the notice and save over a chart they cannot see.
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: FUNNEL });
    const locks = elements((await fallbackOf(zip, MERGED_SLIDES[0] ?? "")) as Element, A_NS, "spLocks");
    expect(locks).toHaveLength(1);
    expect(locks[0]?.getAttribute("noTextEdit")).toBe("1");
  });

  it("puts the notice where the chart is, at the size the chart is", async () => {
    // Taken from the Choice branch's own frame rather than guessed at, so it
    // does not land somewhere else on the slide.
    //
    // The size is checked from INSIDE the notice's `<a:xfrm>`, and the offset
    // with it, because the frame the box is read from holds a second element
    // also called `a:ext` — the extension-list entry in `<p:cNvPr>` carrying a
    // creation id — and it stands earlier in the document. A reader that takes
    // the first `a:ext` by name takes that one, and the notice then has an
    // offset and no size at all: a shape a host cannot draw. The fixture's
    // frame carries the `<a:extLst>` PowerPoint writes precisely so this can
    // go red.
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: FUNNEL });
    const fallback = (await fallbackOf(zip, MERGED_SLIDES[0] ?? "")) as Element;
    const xfrm = elements(fallback, A_NS, "xfrm")[0];
    expect(xfrm, "the notice has no transform at all").toBeDefined();
    const off = children(xfrm as Element, A_NS, "off")[0];
    const ext = children(xfrm as Element, A_NS, "ext")[0];
    expect(off?.getAttribute("x")).toBe("1000000");
    expect(off?.getAttribute("y")).toBe("500000");
    expect(ext?.getAttribute("cx")).toBe("6000000");
    expect(ext?.getAttribute("cy")).toBe("4000000");
    // And nothing else rode along inside it.
    expect(elements(fallback, A_NS, "ext")).toHaveLength(1);
  });

  it("stops relating to the picture, so nothing keeps the old rendering alive", async () => {
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: FUNNEL });
    for (const slide of MERGED_SLIDES) {
      expect(await related(zip, slide, "/image"), `${slide} still points at a picture`).toEqual([]);
    }
    // The TEMPLATE keeps its own: that slide is the author's and is not merged.
    expect(await related(zip, "ppt/slides/slide1.xml", "/image")).toEqual(["ppt/media/chart1.png"]);
  });

  it("keeps the image relationships the slide still names in ways that are not r:embed", async () => {
    // Dropping the fallback picture's relationship means deciding which image
    // relationships a merged copy still needs, and the first version of that
    // decision asked one question: which `<a:blip>` carries which `r:embed`.
    //
    // Two ordinary shapes name an image any other way. A modern PowerPoint
    // ICON is a raster blip carrying the real SVG beside it — `<asvg:svgBlip
    // r:embed>` inside the blip's own extension list, under a SECOND image
    // relationship. A LINKED picture names its relationship with `r:link` and
    // embeds nothing. Both relationships were deleted while the slide still
    // referenced them: a slide naming a relationship that is not there, which
    // is what PowerPoint calls a damaged file.
    const { zip } = await mergeDeck({ paragraphs: [["Cover"]], modernChart: FUNNEL, icons: true });
    for (const slide of MERGED_SLIDES) {
      const body = (await zip.file(slide)?.async("string")) ?? "";
      const named = new Set([...body.matchAll(/r:(?:embed|link)="([^"]+)"/g)].map((m) => m[1] ?? ""));
      expect(named, `${slide} lost the shapes the fixture put on it`).toContain("rId20");
      const dir = slide.slice(0, slide.lastIndexOf("/"));
      const rels = parseXml(
        (await zip.file(`${dir}/_rels/${slide.slice(slide.lastIndexOf("/") + 1)}.rels`)?.async("string")) ?? "",
      );
      // Not "does the id exist" — "does it still lead to a picture". Deleting a
      // relationship frees its ID, and the tag writer takes the next free one:
      // in the build without this fix, `rId21` came back pointing at
      // `ppt/tags/tag1.xml`, so the icon's SVG resolved to a tag part. An id
      // that exists and means something else is worse than one that is gone,
      // and a test asking only whether it is gone cannot see it.
      const images = new Map(
        elements(rels, PKG_REL_NS, "Relationship")
          .filter((r) => (r.getAttribute("Type") ?? "").endsWith("/image"))
          .map((r) => [r.getAttribute("Id") ?? "", r.getAttribute("Target") ?? ""]),
      );
      expect(
        [...named].filter((id) => !images.has(id)),
        `${slide} names a picture through a relationship that is not an image`,
      ).toEqual([]);
    }
  });

  it("leaves a chart with no fallback branch alone", async () => {
    // `mc:Fallback` is optional in the MCE schema and a producer may write
    // none. There is then nothing to replace, and inventing one would put a
    // shape on the slide the author never had.
    const { zip } = await mergeDeck({
      paragraphs: [["Cover"]],
      modernChart: { ...FUNNEL, noFallback: true },
    });
    expect(await fallbackOf(zip, MERGED_SLIDES[0] ?? "")).toBeUndefined();
  });
});
