/**
 * A minimal .pptx built in memory.
 *
 * Hermetic on purpose: every test starts from bytes this file produced, so a
 * failure is about the engine and never about which PowerPoint wrote the
 * fixture. It is small enough to read and structurally real where the engine
 * touches it, which is the parts, the relationships, the content types, the
 * slide id list and the creation id.
 *
 * It is NOT a validity oracle. Nothing here proves PowerPoint would open the
 * file; that is what a round in the real host is for.
 */
import JSZip from "jszip";

const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';

export interface SlideSpec {
  /** Paragraphs, each given as the runs it is split into. Split a placeholder to reproduce the real thing. */
  paragraphs: string[][];
  /** true for stock notes text, or the text itself — a placeholder in it is the point. */
  notes?: boolean | string;
  /**
   * A chart related from this slide.
   *
   * A bare string is its title, which is where this started. The object form
   * reaches the two places a chart's labels REALLY live — its own string cache,
   * and the workbook Excel opens on "Edit Data" — because a merge that fills
   * one and not the other is a deck that contradicts itself.
   */
  chart?: string | ChartSpec;
  /**
   * A MODERN chart related from this slide — the chartEx family.
   *
   * Shaped from a real PowerPoint file (`funnel-pp1.pptx` in LibreOffice's
   * chart test data), because every part of it would have been guessed wrong:
   * the slide carries `<mc:AlternateContent>` with the frame in `Choice` and a
   * PICTURE of the chart in `Fallback`, the labels are `<cx:pt>` inside a
   * `<cx:strDim>`, the values are `<cx:pt>` inside a `<cx:numDim>` — the same
   * element — the series name is `<cx:tx><cx:txData><cx:v>`, and the title is
   * DrawingML with no `<cx:tx>` at all.
   */
  modernChart?: ModernChartSpec;
  /**
   * A modern PowerPoint ICON and a LINKED picture, side by side.
   *
   * The two ways an image relationship is named that are NOT
   * `<a:blip r:embed>`. An icon is a raster blip carrying the real SVG beside
   * it in its own extension list, under a second image relationship; a linked
   * picture names its relationship with `r:link` and embeds nothing. Both are
   * ordinary in a deck somebody made this decade, and both were invisible to
   * the pass that decides which image relationships a merged copy still needs.
   */
  icons?: boolean;
  /**
   * SmartArt related from this slide, given as its node labels.
   *
   * Produces both halves a real one has: `dataN.xml`, the model, and
   * `drawingN.xml`, the laid-out rendering PowerPoint actually displays. A
   * fixture with only the model would pass a merge that leaves every visible
   * label unfilled.
   */
  smartArt?: string[];
  /**
   * Which part owns the `diagramDrawing` relationship.
   *
   * `"data"` — the default and the shape every fixture here used to assume:
   * `ppt/diagrams/_rels/dataN.xml.rels` names the drawing.
   *
   * `"slide"` — **what real PowerPoint writes.** The slide's own relationships
   * name the drawing, `dataN.xml` carries `<dsp:dataModelExt relId="…">`
   * pointing at that id, and there is no `dataN.xml.rels` at all.
   *
   * Both exist because the engine looked only at the data part until
   * 2026-08-28, so a diagram PowerPoint had authored merged its model and left
   * every copy sharing one unmerged rendering — and no fixture could show it,
   * because they were all built the way the reader expected.
   */
  smartArtDrawingOn?: "data" | "slide";
  creationId?: number;
  /**
   * The shape's own `<a:xfrm>`, as XML.
   *
   * Omitted by default because the commonest real shape — a text box on a
   * layout placeholder — states no box and inherits one, and that is the case
   * the picture pass has to degrade for. Supplied when a test needs a RATIO,
   * which cover and contain cannot be computed without.
   */
  box?: string;
}

export interface ModernChartSpec {
  /** The title, as DrawingML inside `<cx:txPr>` — which is where a real funnel keeps it. */
  title?: string;
  /** Category labels, in `<cx:strDim>`. Text. */
  categories?: string[];
  /** The series name, in `<cx:tx><cx:txData><cx:v>`. Text. */
  series?: string;
  /**
   * The plotted values, in `<cx:numDim>`.
   *
   * Text so a test can put a PLACEHOLDER among them: `<cx:pt>` is the same
   * element as a category label's, and the merge must fill one and not the
   * other. Defaults to ordinary numbers.
   */
  values?: string[];
  /** Give it an embedded workbook, with these as its shared strings. */
  workbook?: string[];
  /** Omit the `mc:Fallback` branch, which the MCE schema allows. */
  noFallback?: boolean;
}

export interface ChartSpec {
  /** The chart's title, in DrawingML rich text. */
  title?: string;
  /** Category labels, in the chart's own `<c:strCache>` — what PowerPoint draws. */
  categories?: string[];
  /** Give the chart an embedded workbook, with these as its shared strings. */
  workbook?: string[];
  /**
   * The plotted values, in the chart's `<c:numCache>`.
   *
   * Text rather than numbers so a test can put a PLACEHOLDER here — the one
   * place in a chart a merge must not write, because the content has to parse
   * as a number. Defaults to ordinary numbers.
   */
  values?: string[];
}

/** An `<a:xfrm>` of the given size, for a spec's `box`. */
export function xfrm(cx: number, cy: number): string {
  return `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
}

function escapeText(t: string): string {
  return t.replace(/&/g, "&amp;").replace(/</g, "&lt;");
}

function runs(parts: string[]): string {
  return parts
    .map(
      (t, i) =>
        `<a:r><a:rPr lang="en-US" b="${i === 0 ? 1 : 0}" dirty="0"/><a:t>${t
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")}</a:t></a:r>`,
    )
    .join("");
}

function slideXml(spec: SlideSpec): string {
  const paras = spec.paragraphs.map((p) => `<a:p>${runs(p)}</a:p>`).join("");
  const creation =
    spec.creationId === undefined
      ? ""
      : `<p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">` +
        `<p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="${spec.creationId}"/>` +
        `</p:ext></p:extLst>`;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<p:sld ${P} ${A} ${R}><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${spec.box ?? ""}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody></p:sp>` +
    `${spec.smartArt ? smartArtFrame() : ""}` +
    `${spec.modernChart ? modernChartFrame(spec.modernChart) : ""}` +
    `${spec.icons ? iconShapes() : ""}` +
    `</p:spTree>${creation}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
  );
}

const DGM_NS = 'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"';
const DSP_NS = 'xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"';

/** A placeholder as PowerPoint routinely stores one: split across two runs. */
function splitRuns(text: string): string {
  const half = Math.ceil(text.length / 2);
  return (
    `<a:r><a:rPr lang="en-US"/><a:t>${escapeText(text.slice(0, half))}</a:t></a:r>` +
    `<a:r><a:rPr lang="en-US"/><a:t>${escapeText(text.slice(half))}</a:t></a:r>`
  );
}

/**
 * A chart part, shaped where the engine touches it.
 *
 * The title is DrawingML rich text and the categories are `<c:v>` inside a
 * `<c:strCache>` — two different places, and a real chart's LABELS are the
 * second one. The values are a `<c:numCache>` deliberately: it is the one place
 * in a chart a placeholder may not go, since the text there has to parse as a
 * number, and a test that never has one cannot prove the merge leaves it alone.
 */
function chartXml(spec: ChartSpec): string {
  const title = spec.title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>${splitRuns(spec.title)}</a:p></c:rich></c:tx></c:title>`
    : "";
  const cats = spec.categories ?? [];
  const cat = cats.length
    ? `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${cats.length + 1}</c:f><c:strCache>` +
      `<c:ptCount val="${cats.length}"/>` +
      cats.map((c, i) => `<c:pt idx="${i}"><c:v>${escapeText(c)}</c:v></c:pt>`).join("") +
      `</c:strCache></c:strRef></c:cat>`
    : "";
  const values = spec.values ?? (cats.length ? cats : [""]).map((_, i) => String(i + 1));
  const val =
    `<c:val><c:numRef><c:f>Sheet1!$B$2:$B$${values.length + 1}</c:f><c:numCache>` +
    `<c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>` +
    values.map((v, i) => `<c:pt idx="${i}"><c:v>${escapeText(v)}</c:v></c:pt>`).join("") +
    `</c:numCache></c:numRef></c:val>`;
  const external = spec.workbook ? `<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>` : "";
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ${A} ${R}>` +
    `<c:chart>${title}<c:plotArea><c:barChart><c:ser>` +
    `<c:idx val="0"/><c:order val="0"/>${cat}${val}` +
    `</c:ser></c:barChart></c:plotArea></c:chart>${external}</c:chartSpace>`
  );
}

/**
 * A minimal but real `.xlsx`, as the workbook behind a chart.
 *
 * A package inside the package, which is the whole reason the merge opens it
 * with its own zip reader. The strings go in `sharedStrings.xml`, where Excel
 * puts them, and the sheet references them by index — so a fixture that wrote
 * the text into the sheet instead would pass a merge that never touches shared
 * strings at all.
 */
async function workbookBytes(strings: string[], values: string[] = []): Promise<Uint8Array> {
  // A non-numeric value cell is a shared string too, so it needs an entry in
  // the table and an index the sheet can point at. Appended after the
  // categories, and deduplicated, exactly as Excel would.
  const valueStrings = [...new Set(values.filter((v) => !Number.isFinite(Number(v))))];
  const table = [...strings, ...valueStrings];
  const book = new JSZip();
  book.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` +
      `</Types>`,
  );
  book.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
      `</Relationships>`,
  );
  const S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
  book.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<workbook ${S} ${R}>` +
      `<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  book.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` +
      `</Relationships>`,
  );
  book.file(
    "xl/sharedStrings.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
      `<sst ${S} count="${table.length}" uniqueCount="${table.length}">` +
      // The first string is split into two runs, which is what a shared string
      // edited in Excel looks like and what a per-node search would miss.
      table
        .map((t, i) =>
          i === 0
            ? `<si><r><t>${escapeText(t.slice(0, Math.ceil(t.length / 2)))}</t></r>` +
              `<r><t>${escapeText(t.slice(Math.ceil(t.length / 2)))}</t></r></si>`
            : `<si><t>${escapeText(t)}</t></si>`,
        )
        .join("") +
      `</sst>`,
  );
  // Column A is the categories, as shared strings. Column B is the VALUES, and
  // it is built the way Excel builds it: a number is a numeric cell with no `t`
  // at all, and anything else — a placeholder somebody typed into Edit Data —
  // is an ordinary shared string like any other text in the sheet. That
  // difference is the whole subject of the numeric merge, so a fixture that
  // wrote every cell the same way could not show it.
  const rows = Math.max(strings.length, values.length);
  book.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet ${S}><sheetData>` +
      Array.from({ length: rows }, (_, i) => {
        const r = i + 2;
        const a = i < strings.length ? `<c r="A${r}" t="s"><v>${i}</v></c>` : "";
        const raw = values[i];
        if (raw === undefined) return `<row r="${r}">${a}</row>`;
        const b = Number.isFinite(Number(raw))
          ? `<c r="B${r}"><v>${escapeText(raw)}</v></c>`
          : `<c r="B${r}" t="s"><v>${strings.length + valueStrings.indexOf(raw)}</v></c>`;
        return `<row r="${r}">${a}${b}</row>`;
      }).join("") +
      `</sheetData></worksheet>`,
  );
  return book.generateAsync({ type: "uint8array" });
}

/** The graphic frame that puts SmartArt on a slide. Referenced by rId, never by path. */
function smartArtFrame(): string {
  return (
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Diagram"/>` +
    `<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="3000000"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
    `<dgm:relIds ${DGM_NS} ${R} r:dm="rId4" r:lo="rId5" r:qs="rId6" r:cs="rId7"/>` +
    `</a:graphicData></a:graphic></p:graphicFrame>`
  );
}

/** The SmartArt model: nodes and their text. */
function diagramData(labels: string[]): string {
  const points = labels
    .map(
      (t, i) =>
        `<dgm:pt modelId="{node-${i}}"><dgm:prSet/><dgm:spPr/>` +
        `<dgm:t><a:bodyPr/><a:lstStyle/><a:p>${splitRuns(t)}</a:p></dgm:t></dgm:pt>`,
    )
    .join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<dgm:dataModel ${DGM_NS} ${A} ${R}><dgm:ptLst>${points}</dgm:ptLst><dgm:cxnLst/>` +
    `<dgm:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">` +
    `<dsp:dataModelExt ${DSP_NS} relId="rId1" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>` +
    `</a:ext></dgm:extLst></dgm:dataModel>`
  );
}

/** The SmartArt rendering, which is the half PowerPoint puts on the screen. */
function diagramDrawing(labels: string[]): string {
  const shapes = labels
    .map(
      (t, i) =>
        `<dsp:sp modelId="{node-${i}}"><dsp:spPr/><dsp:txBody><a:bodyPr/><a:lstStyle/>` +
        `<a:p>${splitRuns(t)}</a:p></dsp:txBody></dsp:sp>`,
    )
    .join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<dsp:drawing ${DSP_NS} ${A}><dsp:spTree>${shapes}</dsp:spTree></dsp:drawing>`
  );
}

/** A one-pixel PNG, standing in for the rendering PowerPoint puts in a fallback. */
const PNG_BYTES = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00,
  0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00, 0x0a, 0x49,
  0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00,
  0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

const CX_NS = 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"';
const MC_NS = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';

/**
 * A modern chart part, shaped from the real one.
 *
 * The title is DrawingML inside `<cx:txPr>` and there is no `<cx:tx>` beside
 * it, which is what `funnel-pp1.pptx` does. The series name IS a
 * `<cx:tx><cx:txData><cx:v>`. Category labels and plotted values are both
 * `<cx:pt>`, told apart only by the dim that holds them — the distinction the
 * merge has to make and the one a reader written from the element name gets
 * wrong.
 */
function modernChartXml(spec: ModernChartSpec): string {
  const cats = spec.categories ?? [];
  const values = spec.values ?? cats.map((_, i) => String(i + 1));
  const external = spec.workbook ? `<cx:externalData r:id="rId1" cx:autoUpdate="0"/>` : "";
  const strDim = cats.length
    ? `<cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$${cats.length + 1}</cx:f>` +
      `<cx:lvl ptCount="${cats.length}">` +
      cats.map((c, i) => `<cx:pt idx="${i}">${escapeText(c)}</cx:pt>`).join("") +
      `</cx:lvl></cx:strDim>`
    : "";
  const numDim = values.length
    ? `<cx:numDim type="val"><cx:f>Sheet1!$B$2:$B$${values.length + 1}</cx:f>` +
      `<cx:lvl ptCount="${values.length}" formatCode="General">` +
      values.map((v, i) => `<cx:pt idx="${i}">${escapeText(v)}</cx:pt>`).join("") +
      `</cx:lvl></cx:numDim>`
    : "";
  const series = spec.series
    ? `<cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>${escapeText(spec.series)}</cx:v></cx:txData></cx:tx>`
    : "";
  const title = spec.title
    ? `<cx:title pos="t" align="ctr" overlay="0"><cx:txPr><a:bodyPr/><a:lstStyle/><a:p>${splitRuns(spec.title)}</a:p></cx:txPr></cx:title>`
    : "";
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<cx:chartSpace ${CX_NS} ${A} ${R}>` +
    `<cx:chartData>${external}<cx:data id="0">${strDim}${numDim}</cx:data></cx:chartData>` +
    `<cx:chart>${title}<cx:plotArea><cx:plotAreaRegion>` +
    `<cx:series layoutId="funnel" uniqueId="{00000000-0000-0000-0000-000000000001}">${series}<cx:dataId val="0"/></cx:series>` +
    `</cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`
  );
}

/**
 * The slide markup a modern chart sits in.
 *
 * `mc:Choice` holds the graphic frame; `mc:Fallback` holds a PICTURE of the
 * chart, which is the half this engine replaces. Both branches carry the same
 * shape id, as they do in the real file.
 */
/**
 * An icon and a linked picture, as PowerPoint writes them.
 *
 * The `<a:extLst>` uri is the one Office stamps on an SVG companion, and the
 * shape was copied from a real deck rather than invented: a template this
 * project was handed in August 2026 carries `asvg:svgBlip` on dozens of its
 * layouts.
 */
function iconShapes(): string {
  const pic = (id: number, name: string, fill: string): string =>
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill>${fill}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  return (
    pic(
      90,
      "Icon",
      `<a:blip r:embed="rId20"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">` +
        `<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId21"/>` +
        `</a:ext></a:extLst></a:blip>`,
    ) + pic(91, "Linked", `<a:blip r:link="rId22"/>`)
  );
}

function modernChartFrame(spec: ModernChartSpec): string {
  const box = `<a:off x="1000000" y="500000"/><a:ext cx="6000000" cy="4000000"/>`;
  const fallback = spec.noFallback
    ? ""
    : `<mc:Fallback><p:pic><p:nvPicPr><p:cNvPr id="6" name="Chart 5"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rId8"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm>${box}</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
      `</p:pic></mc:Fallback>`;
  return (
    `<mc:AlternateContent ${MC_NS}>` +
    `<mc:Choice xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" Requires="cx2">` +
    // The `<a:extLst>` is not decoration. PowerPoint writes one in a chart
    // frame's `<p:cNvPr>`, and the `<a:ext uri="{GUID}">` inside it shares its
    // name with the SIZE element in `<p:xfrm>` below while standing earlier in
    // the document — so a reader that searches for the first `a:ext` by name
    // finds a creation id and gives the fallback notice no size at all. Every
    // fixture here was written without one until a deck real PowerPoint had
    // saved showed up carrying it.
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Chart 5">` +
    `<a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">` +
    `<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{460B2BF8-B46D-8C8B-2E20-223AA4200A65}"/>` +
    `</a:ext></a:extLst></p:cNvPr><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm>${box}</p:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">` +
    `<cx:chart ${CX_NS} ${R} r:id="rId7"/></a:graphicData></a:graphic></p:graphicFrame>` +
    `</mc:Choice>${fallback}</mc:AlternateContent>`
  );
}

const TYPE = {
  slide: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
  notes: "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
  chart: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
  chartEx: "application/vnd.ms-office.chartex+xml",
  diagramData: "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
  diagramLayout: "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
  diagramStyle: "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
  diagramColors: "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
  diagramDrawing: "application/vnd.ms-office.drawingml.diagramDrawing+xml",
} as const;

const REL_TYPE = {
  slide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
  notes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
  layout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
  master: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  doc: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  chart: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
  chartEx: "http://schemas.microsoft.com/office/2014/relationships/chartEx",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  package: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
  diagramData: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
  diagramLayout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
  diagramQuickStyle: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
  diagramColors: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
  diagramDrawing: "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
} as const;

/** Build a deck whose slides are exactly the specs given. */
export async function makeDeck(slides: SlideSpec[]): Promise<Uint8Array> {
  const zip = new JSZip();

  const overrides = slides
    .map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="${TYPE.slide}"/>`)
    .concat(
      slides.flatMap((s, i) => [
        ...(s.notes
          ? [`<Override PartName="/ppt/notesSlides/notesSlide${i + 1}.xml" ContentType="${TYPE.notes}"/>`]
          : []),
        ...(s.chart ? [`<Override PartName="/ppt/charts/chart${i + 1}.xml" ContentType="${TYPE.chart}"/>`] : []),
        ...(s.modernChart
          ? [`<Override PartName="/ppt/charts/chartEx${i + 1}.xml" ContentType="${TYPE.chartEx}"/>`]
          : []),
        ...(s.smartArt
          ? [
              `<Override PartName="/ppt/diagrams/data${i + 1}.xml" ContentType="${TYPE.diagramData}"/>`,
              `<Override PartName="/ppt/diagrams/layout${i + 1}.xml" ContentType="${TYPE.diagramLayout}"/>`,
              `<Override PartName="/ppt/diagrams/quickStyle${i + 1}.xml" ContentType="${TYPE.diagramStyle}"/>`,
              `<Override PartName="/ppt/diagrams/colors${i + 1}.xml" ContentType="${TYPE.diagramColors}"/>`,
              `<Override PartName="/ppt/diagrams/drawing${i + 1}.xml" ContentType="${TYPE.diagramDrawing}"/>`,
            ]
          : []),
      ]),
    )
    .join("");

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>` +
      `<Default Extension="png" ContentType="image/png"/>` +
      `<Default Extension="svg" ContentType="image/svg+xml"/>` +
      `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
      `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
      `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
      `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
      `${overrides}</Types>`,
  );

  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
      `<Relationship Id="rId1" Type="${REL_TYPE.doc}" Target="ppt/presentation.xml"/></Relationships>`,
  );

  const sldIds = slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join("");
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:presentation ${P} ${A} ${R}>` +
      `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
      `<p:sldIdLst>${sldIds}</p:sldIdLst>` +
      `<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
  );

  const presRels = slides
    .map((_, i) => `<Relationship Id="rId${i + 2}" Type="${REL_TYPE.slide}" Target="slides/slide${i + 1}.xml"/>`)
    .join("");
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
      `<Relationship Id="rId1" Type="${REL_TYPE.master}" Target="slideMasters/slideMaster1.xml"/>` +
      `${presRels}</Relationships>`,
  );

  zip.file(
    "ppt/slideMasters/slideMaster1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldMaster ${P} ${A} ${R}>` +
      `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
      `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
      `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`,
  );
  zip.file(
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
      `<Relationship Id="rId1" Type="${REL_TYPE.layout}" Target="../slideLayouts/slideLayout1.xml"/>` +
      `<Relationship Id="rId2" Type="${REL_TYPE.theme}" Target="../theme/theme1.xml"/></Relationships>`,
  );

  zip.file(
    "ppt/slideLayouts/slideLayout1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldLayout ${P} ${A} ${R} type="blank">` +
      `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
      `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
  );
  zip.file(
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
      `<Relationship Id="rId1" Type="${REL_TYPE.master}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
  );

  zip.file(
    "ppt/theme/theme1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<a:theme ${A} name="Test"><a:themeElements/></a:theme>`,
  );

  for (const [i, spec] of slides.entries()) {
    const n = i + 1;
    zip.file(`ppt/slides/slide${n}.xml`, slideXml(spec));
    const notesRel = spec.notes
      ? `<Relationship Id="rId2" Type="${REL_TYPE.notes}" Target="../notesSlides/notesSlide${n}.xml"/>`
      : "";
    const chartRel = spec.chart
      ? `<Relationship Id="rId3" Type="${REL_TYPE.chart}" Target="../charts/chart${n}.xml"/>`
      : "";
    const modernRels = spec.modernChart
      ? `<Relationship Id="rId7" Type="${REL_TYPE.chartEx}" Target="../charts/chartEx${n}.xml"/>` +
        (spec.modernChart.noFallback
          ? ""
          : `<Relationship Id="rId8" Type="${REL_TYPE.image}" Target="../media/chart${n}.png"/>`)
      : "";
    const iconRels = spec.icons
      ? `<Relationship Id="rId20" Type="${REL_TYPE.image}" Target="../media/icon${n}.png"/>` +
        `<Relationship Id="rId21" Type="${REL_TYPE.image}" Target="../media/icon${n}.svg"/>` +
        `<Relationship Id="rId22" Type="${REL_TYPE.image}" Target="../media/linked${n}.png"/>`
      : "";
    const diagramRels = spec.smartArt
      ? `<Relationship Id="rId4" Type="${REL_TYPE.diagramData}" Target="../diagrams/data${n}.xml"/>` +
        `<Relationship Id="rId5" Type="${REL_TYPE.diagramLayout}" Target="../diagrams/layout${n}.xml"/>` +
        `<Relationship Id="rId6" Type="${REL_TYPE.diagramQuickStyle}" Target="../diagrams/quickStyle${n}.xml"/>` +
        `<Relationship Id="rId7" Type="${REL_TYPE.diagramColors}" Target="../diagrams/colors${n}.xml"/>` +
        // The PowerPoint shape: the drawing is the SLIDE's relationship.
        (spec.smartArtDrawingOn === "slide"
          ? `<Relationship Id="rId8" Type="${REL_TYPE.diagramDrawing}" Target="../diagrams/drawing${n}.xml"/>`
          : "")
      : "";
    zip.file(
      `ppt/slides/_rels/slide${n}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
        `<Relationship Id="rId1" Type="${REL_TYPE.layout}" Target="../slideLayouts/slideLayout1.xml"/>` +
        `${notesRel}${chartRel}${diagramRels}${modernRels}${iconRels}</Relationships>`,
    );
    if (spec.chart) {
      const chart: ChartSpec = typeof spec.chart === "string" ? { title: spec.chart } : spec.chart;
      zip.file(`ppt/charts/chart${n}.xml`, chartXml(chart));
      if (chart.workbook) {
        zip.file(
          `ppt/embeddings/Microsoft_Excel_Worksheet${n}.xlsx`,
          await workbookBytes(chart.workbook, chart.values ?? []),
        );
        zip.file(
          `ppt/charts/_rels/chart${n}.xml.rels`,
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
            `<Relationship Id="rId1" Type="${REL_TYPE.package}" Target="../embeddings/Microsoft_Excel_Worksheet${n}.xlsx"/>` +
            `</Relationships>`,
        );
      }
    }
    if (spec.icons) {
      // Bytes enough to be a part; nothing reads them, the relationships are
      // the whole subject.
      zip.file(`ppt/media/icon${n}.png`, PNG_BYTES);
      zip.file(`ppt/media/icon${n}.svg`, new TextEncoder().encode(`<svg xmlns="http://www.w3.org/2000/svg"/>`));
      zip.file(`ppt/media/linked${n}.png`, PNG_BYTES);
    }
    if (spec.modernChart) {
      zip.file(`ppt/charts/chartEx${n}.xml`, modernChartXml(spec.modernChart));
      if (spec.modernChart.workbook) {
        // The VALUES go in as well, exactly as the classic branch above does
        // it. Without them the sheet had a category column and no value column
        // at all, so a placeholder in `<cx:numDim>` named a cell that did not
        // exist and the numeric merge could not be reached — the fixture, not
        // the engine, was what made a modern chart's values unmergeable.
        zip.file(
          `ppt/embeddings/Modern_Worksheet${n}.xlsx`,
          await workbookBytes(spec.modernChart.workbook, spec.modernChart.values ?? []),
        );
        zip.file(
          `ppt/charts/_rels/chartEx${n}.xml.rels`,
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
            `<Relationship Id="rId1" Type="${REL_TYPE.package}" Target="../embeddings/Modern_Worksheet${n}.xlsx"/>` +
            `</Relationships>`,
        );
      }
      if (!spec.modernChart.noFallback) {
        // A stand-in for the RENDERED PNG PowerPoint puts in the fallback. Its
        // bytes do not matter; that a merged copy stops pointing at it does.
        zip.file(`ppt/media/chart${n}.png`, PNG_BYTES);
      }
    }
    if (spec.smartArt) {
      zip.file(`ppt/diagrams/data${n}.xml`, diagramData(spec.smartArt));
      zip.file(`ppt/diagrams/drawing${n}.xml`, diagramDrawing(spec.smartArt));
      // Layout, quick style and colours are read-only styling and stay shared
      // between every copy — which is what the cloning has to get right, so the
      // fixture has to have them.
      for (const [file, root] of [
        [`layout${n}`, "layoutDef"],
        [`quickStyle${n}`, "styleDef"],
        [`colors${n}`, "colorsDef"],
      ] as const) {
        zip.file(
          `ppt/diagrams/${file}.xml`,
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<dgm:${root} ${DGM_NS} ${A} uniqueId="${file}"/>`,
        );
      }
      // Where the drawing hangs is the whole point of this option. In the
      // "slide" shape there is deliberately NO dataN.xml.rels — that absence is
      // what real PowerPoint writes, and what the engine used to walk into.
      if (spec.smartArtDrawingOn !== "slide") {
        zip.file(
          `ppt/diagrams/_rels/data${n}.xml.rels`,
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
            `<Relationship Id="rId1" Type="${REL_TYPE.diagramDrawing}" Target="drawing${n}.xml"/></Relationships>`,
        );
      }
    }
    if (spec.notes) {
      zip.file(
        `ppt/notesSlides/notesSlide${n}.xml`,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:notes ${P} ${A} ${R}>` +
          `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${typeof spec.notes === "string" ? escapeText(spec.notes) : `notes for slide ${n}`}</a:t></a:r></a:p></p:txBody>` +
          `</p:sp></p:spTree></p:cSld></p:notes>`,
      );
      zip.file(
        `ppt/notesSlides/_rels/notesSlide${n}.xml.rels`,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
          `<Relationship Id="rId1" Type="${REL_TYPE.slide}" Target="../slides/slide${n}.xml"/></Relationships>`,
      );
    }
  }

  return zip.generateAsync({ type: "uint8array" });
}
