/**
 * Give a cloned slide its own charts and SmartArt.
 *
 * `cloneSlide` copies a slide's relationships wholesale, which is right for
 * everything read-only — the layout, the master, the theme, media — and wrong
 * for anything the merge is about to WRITE into. A chart and a SmartArt diagram
 * are both content parts hanging off the slide, so without this every copy of
 * the template points at the template's own `chart1.xml`: merging into it would
 * write the last record's values once, and all 240 slides would show them.
 *
 * That is the third time this shape has been found here. Notes pages shared the
 * same way until they were cloned per copy; comments shared the same way and are
 * dropped. The rule the three of them make: **a part a merge writes into may
 * never be shared by two slides.**
 *
 * What stays shared is the other half of the rule, and it is not an oversight.
 * A chart's colour and style parts, and a diagram's layout, quick style and
 * colours, are read-only styling — nothing in a merge touches them, and copying
 * them per record would multiply a template's styling by the row count for no
 * change in what anybody sees.
 *
 * The EMBEDDED WORKBOOK is copied, because it is written into. A chart's labels
 * live twice: in the chart's own string cache, which is what PowerPoint draws,
 * and in the workbook behind it, which is what Excel opens on "Edit Data" — and
 * closing that Excel refreshes the cache from the workbook. Merge the cache
 * alone and the deck is right until somebody clicks the button, at which point
 * the merged labels revert to `{{Region}}` in front of them.
 */
import { Pkg } from "./pkg.js";
import { PKG_REL_NS, elements } from "./xml.js";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MS_REL = "http://schemas.microsoft.com/office/2007/relationships";

export const CHART_REL_TYPE = `${REL}/chart`;
export const DIAGRAM_DATA_REL_TYPE = `${REL}/diagramData`;
export const DIAGRAM_DRAWING_REL_TYPE = `${MS_REL}/diagramDrawing`;
/** The embedded workbook behind a chart. Declared as a package, not as a part. */
export const PACKAGE_REL_TYPE = `${REL}/package`;

const CHART_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const DIAGRAM_DATA_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml";
const DIAGRAM_DRAWING_TYPE = "application/vnd.ms-office.drawingml.diagramDrawing+xml";
const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/** Resolve a relationship target against the part that owns it. */
function resolve(ownerPart: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = ownerPart.slice(0, ownerPart.lastIndexOf("/")).split("/");
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

/** The relationships of a part, as elements, or an empty list when it has none. */
async function relsOf(pkg: Pkg, part: string): Promise<Element[]> {
  const path = Pkg.relsPathFor(part);
  if (!pkg.has(path)) return [];
  return elements(await pkg.doc(path), PKG_REL_NS, "Relationship");
}

/**
 * Copy a part and its own relationships, and declare its content type.
 *
 * The rels go with it because the copy has to keep pointing at whatever the
 * original pointed at — a chart at its style parts and its workbook, a diagram
 * at its layout and its drawing. Callers then repoint the few of those that are
 * per-copy content.
 */
async function copyWithRels(pkg: Pkg, from: string, to: string, contentType: string): Promise<void> {
  await pkg.copyPart(from, to);
  const rels = Pkg.relsPathFor(from);
  if (pkg.has(rels)) await pkg.copyPart(rels, Pkg.relsPathFor(to));
  await pkg.addContentTypeOverride(`/${to}`, contentType);
}

/**
 * Point one relationship at a new target, by id.
 *
 * The `rId` is left exactly as it was, which is what makes this safe: the id is
 * referenced from INSIDE the part that owns the relationship — `r:embed` on a
 * blip, `r:dm` on a SmartArt frame, `r:id` on a chart's graphic frame — and
 * renumbering it would need every one of those rewritten too.
 */
async function repoint(pkg: Pkg, ownerPart: string, rId: string, target: string): Promise<void> {
  for (const rel of await relsOf(pkg, ownerPart)) {
    if (rel.getAttribute("Id") === rId) rel.setAttribute("Target", target);
  }
}

/**
 * Give this slide its own copy of every part a merge writes into.
 *
 * Called on the CLONE, after `cloneSlide` has copied the rels. The slide's own
 * markup is untouched: every reference in it is by `rId`, and the ids do not
 * change — only what they point at.
 */
export async function cloneSlideGraphics(pkg: Pkg, slidePath: string): Promise<void> {
  for (const rel of await relsOf(pkg, slidePath)) {
    if ((rel.getAttribute("TargetMode") ?? "") === "External") continue;
    const type = rel.getAttribute("Type") ?? "";
    const target = rel.getAttribute("Target");
    const rId = rel.getAttribute("Id");
    if (!target || !rId) continue;
    const source = resolve(slidePath, target);
    if (!pkg.has(source)) continue;

    if (type === CHART_REL_TYPE) {
      const n = pkg.nextNumber("ppt/charts/chart");
      const path = `ppt/charts/chart${n}.xml`;
      await copyWithRels(pkg, source, path, CHART_TYPE);
      await cloneChartWorkbook(pkg, path);
      await repoint(pkg, slidePath, rId, `../charts/chart${n}.xml`);
    } else if (type === DIAGRAM_DATA_REL_TYPE) {
      const n = pkg.nextNumber("ppt/diagrams/data");
      const path = `ppt/diagrams/data${n}.xml`;
      await copyWithRels(pkg, source, path, DIAGRAM_DATA_TYPE);
      await cloneDiagramDrawing(pkg, path);
      await repoint(pkg, slidePath, rId, `../diagrams/data${n}.xml`);
    }
  }
}

/**
 * Give a copied chart its own workbook.
 *
 * Named from the EMBEDDINGS, never from the chart. Part names in a package are
 * arbitrary and the two sequences drift the moment anything is deleted, so
 * taking the chart's number lands on an embedding that is already there — which
 * `copyPart` overwrites silently, leaving two charts editing one workbook. The
 * notes page cost this project exactly that mistake once already.
 *
 * A chart with no workbook is ordinary rather than broken: a chart pasted as a
 * picture-of-data, or one whose external data was stripped, simply has no
 * `package` relationship, and its cache is all there is.
 */
async function cloneChartWorkbook(pkg: Pkg, chartPath: string): Promise<void> {
  for (const rel of await relsOf(pkg, chartPath)) {
    if (rel.getAttribute("Type") !== PACKAGE_REL_TYPE) continue;
    const target = rel.getAttribute("Target");
    if (!target || (rel.getAttribute("TargetMode") ?? "") === "External") continue;
    const source = resolve(chartPath, target);
    if (!pkg.has(source)) continue;
    const extension = source.slice(source.lastIndexOf(".") + 1);
    const n = pkg.nextNumber("ppt/embeddings/workbook", `.${extension}`);
    const path = `ppt/embeddings/workbook${n}.${extension}`;
    await pkg.copyPart(source, path);
    // A Default rather than an Override: an embedded workbook is declared by
    // its extension the way media is, and the template's own declaration is
    // already there in the ordinary case. Adding it is for the case where the
    // template declared the original with an Override and this copy's name is
    // not covered by it.
    if (extension.toLowerCase() === "xlsx") await pkg.addContentTypeDefault(extension, XLSX_TYPE);
    rel.setAttribute("Target", `../embeddings/workbook${n}.${extension}`);
  }
}

/**
 * Give a copied SmartArt its own drawing.
 *
 * The drawing is the half PowerPoint actually shows. `dataN.xml` is the model —
 * nodes, text, connections — and `drawingN.xml` is the laid-out rendering of it
 * that the host displays without re-running the layout engine. Merge the model
 * alone and the deck shows the placeholders, because the drawing is what is on
 * the screen.
 *
 * It hangs off the DATA part, not off the slide, which is why this runs here
 * rather than in the loop above: a slide's own relationships name the data, the
 * layout, the quick style and the colours, and never the drawing.
 *
 * `<dsp:dataModelExt relId="…">` inside the data part names this relationship
 * by id. The id is unchanged — only the target moves — so that reference stays
 * correct without the data part being rewritten.
 */
async function cloneDiagramDrawing(pkg: Pkg, dataPath: string): Promise<void> {
  for (const rel of await relsOf(pkg, dataPath)) {
    if (rel.getAttribute("Type") !== DIAGRAM_DRAWING_REL_TYPE) continue;
    const target = rel.getAttribute("Target");
    if (!target) continue;
    const source = resolve(dataPath, target);
    if (!pkg.has(source)) continue;
    const n = pkg.nextNumber("ppt/diagrams/drawing");
    const path = `ppt/diagrams/drawing${n}.xml`;
    await pkg.copyPart(source, path);
    await pkg.addContentTypeOverride(`/${path}`, DIAGRAM_DRAWING_TYPE);
    rel.setAttribute("Target", `drawing${n}.xml`);
  }
}

/**
 * Every part of this slide's whose text the merge fills, beyond the slide itself.
 *
 * Charts, SmartArt models and the SmartArt drawings behind them, in the order
 * the slide relates to them. The workbook is NOT here: it is a package inside
 * the package, so it is opened and merged by its own pass rather than as an XML
 * part of this one.
 */
export async function graphicPartsOf(pkg: Pkg, slidePath: string): Promise<string[]> {
  const out: string[] = [];
  for (const rel of await relsOf(pkg, slidePath)) {
    const type = rel.getAttribute("Type") ?? "";
    const target = rel.getAttribute("Target");
    if (!target || (rel.getAttribute("TargetMode") ?? "") === "External") continue;
    if (type !== CHART_REL_TYPE && type !== DIAGRAM_DATA_REL_TYPE) continue;
    const path = resolve(slidePath, target);
    if (!pkg.has(path) || out.includes(path)) continue;
    out.push(path);
    if (type !== DIAGRAM_DATA_REL_TYPE) continue;
    for (const drawing of await relsOf(pkg, path)) {
      if (drawing.getAttribute("Type") !== DIAGRAM_DRAWING_REL_TYPE) continue;
      const drawingTarget = drawing.getAttribute("Target");
      if (!drawingTarget) continue;
      const drawingPath = resolve(path, drawingTarget);
      if (pkg.has(drawingPath) && !out.includes(drawingPath)) out.push(drawingPath);
    }
  }
  return out;
}

/**
 * The workbooks behind this slide's charts, as package paths.
 *
 * Separate from `graphicPartsOf` because these are not XML parts of this
 * package at all — each is a whole `.xlsx` held as one binary part, opened by
 * the merge pass with its own zip reader.
 */
export async function chartWorkbooksOf(pkg: Pkg, slidePath: string): Promise<string[]> {
  const out: string[] = [];
  for (const chart of await graphicPartsOf(pkg, slidePath)) {
    if (!chart.startsWith("ppt/charts/")) continue;
    for (const rel of await relsOf(pkg, chart)) {
      if (rel.getAttribute("Type") !== PACKAGE_REL_TYPE) continue;
      if ((rel.getAttribute("TargetMode") ?? "") === "External") continue;
      const target = rel.getAttribute("Target");
      if (!target) continue;
      const path = resolve(chart, target);
      if (pkg.has(path) && !out.includes(path)) out.push(path);
    }
  }
  return out;
}
