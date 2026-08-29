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
import { Pkg, resolveTarget as resolve } from "./pkg.js";
import { REL_TYPE } from "./parts.js";
import { A_NS, MC_NS, PKG_REL_NS, P_NS, R_NS, child, element, elements } from "./xml.js";

/** The embedded workbook behind a chart. Declared as a package, not as a part. */

const CHART_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const CHARTEX_TYPE = "application/vnd.ms-office.chartex+xml";
const DIAGRAM_DATA_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml";
const DIAGRAM_DRAWING_TYPE = "application/vnd.ms-office.drawingml.diagramDrawing+xml";
const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

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

    if (type === REL_TYPE.chart) {
      const n = pkg.nextNumber("ppt/charts/chart");
      const path = `ppt/charts/chart${n}.xml`;
      await copyWithRels(pkg, source, path, CHART_TYPE);
      await cloneChartWorkbook(pkg, path);
      await repoint(pkg, slidePath, rId, `../charts/chart${n}.xml`);
    } else if (type === REL_TYPE.chartEx) {
      // A modern chart takes the same three steps as a classic one, and its
      // workbook takes a fourth that costs nothing: the embedding hangs off it
      // under the ORDINARY `package` relationship, so `cloneChartWorkbook`
      // needed no change at all. Numbered in its own `chartEx` sequence, which
      // is what PowerPoint names them.
      const n = pkg.nextNumber("ppt/charts/chartEx");
      const path = `ppt/charts/chartEx${n}.xml`;
      await copyWithRels(pkg, source, path, CHARTEX_TYPE);
      await cloneChartWorkbook(pkg, path);
      await repoint(pkg, slidePath, rId, `../charts/chartEx${n}.xml`);
    } else if (type === REL_TYPE.diagramData) {
      const n = pkg.nextNumber("ppt/diagrams/data");
      const path = `ppt/diagrams/data${n}.xml`;
      await copyWithRels(pkg, source, path, DIAGRAM_DATA_TYPE);
      await cloneDiagramDrawing(pkg, path);
      await repoint(pkg, slidePath, rId, `../diagrams/data${n}.xml`);
    } else if (type === REL_TYPE.diagramDrawing) {
      // The shape real PowerPoint writes: the drawing hangs off the SLIDE, and
      // the data part has no relationships of its own. Handled here rather than
      // inside `cloneDiagramDrawing` because there is nothing to reach it from —
      // that function walks the data part, which in this shape is empty.
      //
      // Order does not matter. This and the data branch act on two independent
      // relationships of the same slide, and each repoints only its own.
      await repoint(pkg, slidePath, rId, `../diagrams/${await copyDiagramDrawing(pkg, source)}`);
    }
  }
  // After the parts, because it reads the slide's markup and its relationships
  // and both are settled by now. Cheap on every other deck: a slide with no
  // modern chart has no `<mc:AlternateContent>` to walk.
  await replaceModernChartFallbacks(pkg, slidePath);
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
    if (rel.getAttribute("Type") !== REL_TYPE.package) continue;
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
 * It can hang off EITHER the data part or the slide, and which one is not a
 * detail: it decides whether this function finds anything at all.
 *
 * This used to say the slide's relationships "never" name the drawing, and only
 * looked at the data part. Real PowerPoint does the opposite. A Basic Process
 * diagram inserted by PowerPoint on 2026-08-28 wrote
 * `ppt/slides/_rels/slide3.xml.rels` holding the `diagramDrawing` relationship,
 * `data1.xml` holding `<dsp:dataModelExt relId="rId6"/>` naming it — and **no
 * `ppt/diagrams/_rels/data1.xml.rels` at all**. So the loop ran over an empty
 * relationship set, every branch `continue`d, and the function returned having
 * done nothing, silently. Three merged copies shared one unmerged drawing.
 *
 * `test/fixtures/deck.ts` built its fixtures the other way and said so, so the
 * suite agreed with the reader and neither had met a diagram PowerPoint wrote.
 *
 * Both shapes are handled now, because the fixture's shape is legal too and
 * removing it would be a guess about producers this project has not seen.
 *
 * `<dsp:dataModelExt relId="…">` names this relationship by id. The id is left
 * alone — only the target moves — so that reference stays correct without the
 * data part being rewritten, whichever part owns the relationship.
 */
async function cloneDiagramDrawing(pkg: Pkg, dataPath: string): Promise<void> {
  for (const rel of await relsOf(pkg, dataPath)) {
    if (rel.getAttribute("Type") !== REL_TYPE.diagramDrawing) continue;
    const target = rel.getAttribute("Target");
    if (!target) continue;
    const source = resolve(dataPath, target);
    if (!pkg.has(source)) continue;
    rel.setAttribute("Target", `${await copyDiagramDrawing(pkg, source)}`);
  }
}

/**
 * Copy one drawing part, and answer its new file name.
 *
 * The name rather than the path, because both callers write it into a
 * relationship that already sits in `ppt/diagrams/`.
 *
 * `copyWithRels`, not `copyPart`: a drawing can carry relationships of its own —
 * a picture inside a SmartArt node is an image relationship from this part — and
 * a copy without them points at nothing.
 */
async function copyDiagramDrawing(pkg: Pkg, source: string): Promise<string> {
  const n = pkg.nextNumber("ppt/diagrams/drawing");
  const path = `ppt/diagrams/drawing${n}.xml`;
  await copyWithRels(pkg, source, path, DIAGRAM_DRAWING_TYPE);
  return `drawing${n}.xml`;
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
    // REL_TYPE.diagramDrawing is here because PowerPoint hangs the drawing off
    // the SLIDE. Without it the drawing was cloned per copy but never FILLED,
    // so each merged slide got its own rendering still reading `{{Region}}`.
    if (
      type !== REL_TYPE.chart &&
      type !== REL_TYPE.chartEx &&
      type !== REL_TYPE.diagramData &&
      type !== REL_TYPE.diagramDrawing
    ) {
      continue;
    }
    const path = resolve(slidePath, target);
    if (!pkg.has(path) || out.includes(path)) continue;
    out.push(path);
    // Only a data part has a drawing hanging off it to chase. A drawing found
    // on the slide is already in `out`, and a chart has no such relationship.
    if (type !== REL_TYPE.diagramData) continue;
    for (const drawing of await relsOf(pkg, path)) {
      if (drawing.getAttribute("Type") !== REL_TYPE.diagramDrawing) continue;
      const drawingTarget = drawing.getAttribute("Target");
      if (!drawingTarget) continue;
      const drawingPath = resolve(path, drawingTarget);
      if (pkg.has(drawingPath) && !out.includes(drawingPath)) out.push(drawingPath);
    }
  }
  return out;
}

/**
 * The workbook behind ONE chart, or undefined.
 *
 * `chartWorkbooksOf` answers a slide's whole set, which is what the text pass
 * wants. The numeric pass needs the pairing instead: a value lives in one
 * chart's cache and in one workbook's cell, and filling them from different
 * charts is exactly the mix-up a count would not catch.
 */
export async function packagesOfChart(pkg: Pkg, chartPath: string): Promise<string[]> {
  if (!chartPath.startsWith("ppt/charts/")) return [];
  const out: string[] = [];
  for (const rel of await relsOf(pkg, chartPath)) {
    if (rel.getAttribute("Type") !== REL_TYPE.package) continue;
    if ((rel.getAttribute("TargetMode") ?? "") === "External") continue;
    const target = rel.getAttribute("Target");
    if (!target) continue;
    const path = resolve(chartPath, target);
    if (pkg.has(path) && !out.includes(path)) out.push(path);
  }
  return out;
}

/**
 * What a host too old for a modern chart is shown instead.
 *
 * PowerPoint writes the fallback branch as a PICTURE — a PNG of the chart as it
 * drew it — because it has a renderer and regenerates that picture whenever it
 * saves. This project has neither, and merging cannot produce one: the copy's
 * chart says something different from the template's the moment it is filled.
 *
 * So the picture is replaced rather than kept, and keeping it is the option
 * that had to be refused. On a mail merge the template's rendering is not
 * merely stale, it is ANOTHER RECIPIENT'S FIGURES under this recipient's name,
 * in the file that gets sent out. That is the same reasoning that keeps a
 * placeholder with no column visible instead of blanking it: a thing that looks
 * finished and is not is worse than a thing that says what it is.
 *
 * Dropping the branch entirely is legal — `mc:Fallback` is optional — and it is
 * silent: an old host shows a hole and no reason for it. The shape below costs
 * the same to write and says why, and its shape is known-good because
 * LibreOffice ships one very like it for the same reason
 * (`writeChartexAlternateContent`, `oox/source/export/chartexport.cxx`).
 *
 * Nothing on a current host ever sees this. `mc:Choice` wins wherever the
 * chartex namespace is understood, so the audience is PowerPoint 2013 and
 * earlier plus third-party viewers.
 */
const FALLBACK_TEXT =
  "This chart needs a newer version of PowerPoint. Its data is in this file — open the deck in a version that supports modern charts to see it.";

/**
 * Swap every modern chart's fallback picture on this slide for the notice.
 *
 * Returns how many branches were replaced. Zero is the ordinary answer: a deck
 * with no modern chart has no `<mc:AlternateContent>` at all, and one whose
 * author wrote no fallback has nothing to replace.
 */
export async function replaceModernChartFallbacks(pkg: Pkg, slidePath: string): Promise<number> {
  const doc = await pkg.doc(slidePath);
  let replaced = 0;

  for (const alternate of elements(doc, MC_NS, "AlternateContent")) {
    const choice = child(alternate, MC_NS, "Choice");
    const fallback = child(alternate, MC_NS, "Fallback");
    if (!choice || !fallback) continue;
    // Matched on the graphicData URI rather than on the `Requires` token: the
    // token is cx1, cx2 or cx4 depending on which dated namespace the layout
    // came from, and a reader keying on one misses the other two.
    const isChart = elements(choice, A_NS, "graphicData").some((d) =>
      (d.getAttribute("uri") ?? "").endsWith("/chartex"),
    );
    if (!isChart) continue;

    // The graphic frame's own box, so the notice lands where the chart is
    // rather than at a guessed position — read out of `<p:xfrm>` and nowhere
    // else.
    //
    // `a:ext` is TWO elements with one name. It is a size, `<a:ext cx cy>`
    // inside an `xfrm`, and it is an extension-list entry, `<a:ext uri="{GUID}">`
    // inside an `<a:extLst>`, which is where a producer hangs its own markup.
    // PowerPoint writes an `<a:extLst>` in the frame's `<p:cNvPr>` carrying a
    // creation id, so a descendant search for the first `a:ext` in the Choice
    // branch finds THAT — before the `<p:xfrm>` it was looking for — and the
    // notice was given an `<a:xfrm>` holding an offset and a creation id in
    // place of a size. A shape with no `cx`/`cy` is a shape a host cannot draw.
    //
    // Every fixture here was written without an `extLst`, so the suite agreed
    // with the reader; the first chart real PowerPoint wrote had one. Same
    // overloaded-element trap as `<cx:pt>`, which is a label in a `strDim` and a
    // value in a `numDim`: scope by the parent, never by the name.
    const xfrm = element(choice, P_NS, "xfrm");
    const off = xfrm && child(xfrm, A_NS, "off");
    const ext = xfrm && child(xfrm, A_NS, "ext");
    while (fallback.firstChild) fallback.removeChild(fallback.firstChild);
    fallback.appendChild(notice(doc, off, ext));
    replaced++;
  }

  if (replaced > 0) await dropUnusedImageRels(pkg, slidePath, doc);
  return replaced;
}

/** The notice shape itself: a bordered white box, locked against editing, carrying the sentence. */
function notice(doc: Document, off: Element | undefined, ext: Element | undefined): Element {
  const p = (local: string): Element => doc.createElementNS(P_NS, `p:${local}`);
  const a = (local: string): Element => doc.createElementNS(A_NS, `a:${local}`);

  const sp = p("sp");
  const nvSpPr = p("nvSpPr");
  const cNvPr = p("cNvPr");
  // Id 0 and an empty name, as LibreOffice writes them. A shape inside a
  // fallback is never addressed by anything, and inventing an id risks
  // colliding with the Choice branch's own — which carries the SAME id in every
  // real file, because the two branches are one shape.
  cNvPr.setAttribute("id", "0");
  cNvPr.setAttribute("name", "");
  const cNvSpPr = p("cNvSpPr");
  const locks = a("spLocks");
  // So that somebody on an old host cannot type into the notice and save over
  // a chart they cannot see.
  locks.setAttribute("noTextEdit", "1");
  cNvSpPr.appendChild(locks);
  // `appendChild` one at a time: xmldom implements the DOM Level 3 core and
  // not `ParentNode.append`, which is a living-standard convenience.
  for (const node of [cNvPr, cNvSpPr, p("nvPr")]) nvSpPr.appendChild(node);

  const spPr = p("spPr");
  if (off && ext) {
    const xfrm = a("xfrm");
    // Cloned from the Choice branch, so the notice lands exactly where the
    // chart is rather than at a guessed position.
    for (const node of [off, ext]) xfrm.appendChild(node.cloneNode(true));
    spPr.appendChild(xfrm);
  }
  const geom = a("prstGeom");
  geom.setAttribute("prst", "rect");
  geom.appendChild(a("avLst"));
  const fill = a("solidFill");
  const white = a("prstClr");
  white.setAttribute("val", "white");
  fill.appendChild(white);
  const ln = a("ln");
  ln.setAttribute("w", "12700");
  const lnFill = a("solidFill");
  const grey = a("prstClr");
  grey.setAttribute("val", "gray");
  lnFill.appendChild(grey);
  ln.appendChild(lnFill);
  for (const node of [geom, fill, ln]) spPr.appendChild(node);

  const txBody = p("txBody");
  const bodyPr = a("bodyPr");
  bodyPr.setAttribute("vertOverflow", "clip");
  bodyPr.setAttribute("horzOverflow", "clip");
  bodyPr.setAttribute("wrap", "square");
  const para = a("p");
  const run = a("r");
  const rPr = a("rPr");
  rPr.setAttribute("lang", "en-US");
  rPr.setAttribute("sz", "1100");
  const t = a("t");
  t.textContent = FALLBACK_TEXT;
  for (const node of [rPr, t]) run.appendChild(node);
  para.appendChild(run);
  for (const node of [bodyPr, a("lstStyle"), para]) txBody.appendChild(node);

  for (const node of [nvSpPr, spPr, txBody]) sp.appendChild(node);
  return sp;
}

/**
 * Drop relationships to images the slide no longer references.
 *
 * The fallback pictures were the only thing pointing at them, so leaving the
 * relationships behind would keep a rendering of the TEMPLATE's data alive in
 * the package — reachable by anything that walks relationships rather than
 * markup, and swept by nothing, because the sweep only removes a part no
 * relationship names.
 *
 * Counted from the slide's own markup, never assumed: a template can perfectly
 * well use one image as a fallback picture and again as a logo, and removing
 * the relationship then breaks the logo.
 */
async function dropUnusedImageRels(pkg: Pkg, slidePath: string, doc: Document): Promise<void> {
  const relsPath = Pkg.relsPathFor(slidePath);
  if (!pkg.has(relsPath)) return;
  const used = relationshipIdsIn(doc);
  for (const rel of await relsOf(pkg, slidePath)) {
    if (rel.getAttribute("Type") !== REL_TYPE.image) continue;
    const id = rel.getAttribute("Id");
    if (id && !used.has(id)) rel.parentNode?.removeChild(rel);
  }
}

/**
 * Every relationship id this part's markup names, whatever names it.
 *
 * The question is "may this relationship go", and the only safe way to ask it
 * is of the WHOLE document: any attribute in the relationship namespace is a
 * reference, and a reference means the relationship has to stay.
 *
 * This began as `a:blip/@r:embed`, which is where a picture's image sits and is
 * not the only place an image id appears. Two of them are ordinary:
 *
 * - `<asvg:svgBlip r:embed="…">`, inside the blip's own extension list. That is
 *   how PowerPoint stores an ICON — a raster fallback in the blip and the real
 *   SVG beside it, under a SECOND image relationship. Icons are everywhere in a
 *   modern deck.
 * - `<a:blip r:link="…">`, a picture LINKED rather than embedded.
 *
 * Neither is an `r:embed` on an `a:blip`, so both relationships were dropped
 * from a merged copy whose slide still referenced them — a slide naming a
 * relationship that is not there, which is precisely what PowerPoint calls a
 * damaged file. It needed a modern chart on the slide to fire, because that is
 * what runs this pass at all, but nothing about the shape is exotic.
 *
 * Reading every `r:`-namespaced attribute rather than a list of the ones known
 * today is the conservative direction: an unknown reference keeps a
 * relationship that could have gone, where a missed one breaks the file. The
 * fallback picture is still dropped, because the replacement takes its `<p:pic>`
 * out of the document first and its id is then named by nothing.
 *
 * Matched by namespace OR by prefix. `getAttributeNS` was already paired with a
 * `getAttribute("r:embed")` fallback here for the same reason: a document that
 * came out of a host may not carry the namespace where a reader expects it.
 */
function relationshipIdsIn(doc: Document): Set<string> {
  const used = new Set<string>();
  const walk = (node: Element): void => {
    const attrs = node.attributes;
    for (let i = 0; i < (attrs?.length ?? 0); i++) {
      const attr = attrs?.item(i);
      if (!attr) continue;
      if (attr.namespaceURI === R_NS || attr.name.startsWith("r:")) used.add(attr.value);
    }
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType === 1) walk(child as Element);
    }
  };
  if (doc.documentElement) walk(doc.documentElement);
  return used;
}
