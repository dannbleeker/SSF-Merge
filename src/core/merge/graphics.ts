/**
 * Fill the placeholders inside a slide's charts and SmartArt.
 *
 * Their text is DrawingML — the same `<a:p>` and `<a:t>` a slide holds — so the
 * reader that finds a placeholder split across runs on a slide finds one split
 * across runs in a chart title too. What is different is WHERE ELSE the same
 * string is kept, and every one of those copies has to move together or the
 * deck contradicts itself:
 *
 * - a chart's category and series labels are in its own `<c:strCache>`, which
 *   is what PowerPoint draws, AND in the workbook behind it, which is what
 *   Excel opens on "Edit Data". Closing that Excel refreshes the cache from the
 *   workbook, so a merge that fills only the cache is undone by a click.
 * - a SmartArt node's text is in `dataN.xml`, the model, AND in `drawingN.xml`,
 *   the laid-out rendering the host puts on the screen. The drawing is the half
 *   anybody sees.
 *
 * `cloneSlideGraphics` has already given this slide its own copies of all four,
 * so writing here writes into one record's chart and nobody else's.
 */
import JSZip from "jszip";
import { Pkg } from "../pptx/pkg.js";

import { parseXml, serializeXml } from "../pptx/xml.js";
import { mergeDocument, type Resolve } from "./text.js";
import { emptyNumberOutcome, mergeChartNumbers, tallyNumbers, type NumberOutcome } from "./numbers.js";
import { graphicsOf, workbooksOf, type FieldSite } from "./sites.js";

export interface GraphicOutcome {
  /** Text groups filled in chart and SmartArt parts. */
  merged: number;
  /** Workbooks behind a chart whose strings were filled too. */
  workbooks: number;
  /** Workbooks the merge could not open, by part path. Reported, never thrown on. */
  unreadable: string[];
  /** Chart VALUES filled from the row, and the ones that refused to be numbers. */
  numbers: NumberOutcome;
}

export function emptyGraphicOutcome(): GraphicOutcome {
  return { merged: 0, workbooks: 0, unreadable: [], numbers: emptyNumberOutcome() };
}

/**
 * The parts of a workbook that hold text a merge can fill.
 *
 * `sharedStrings.xml` is where Excel keeps every string a cell shows, and it is
 * where a chart's own labels come from in practice. A worksheet is read too
 * because a cell may hold its string INLINE (`<is><t>`), which is what a
 * generator that never built a shared-string table produces — and a chart
 * written by a tool rather than by Excel is exactly the case that does that.
 *
 * Nothing else is touched. A workbook holds styles, a calc chain and a
 * definition of the range the chart reads; rewriting any of those would change
 * what the chart plots, where this pass only changes what it says.
 */
const WORKBOOK_TEXT = /^xl\/(sharedStrings\.xml|worksheets\/sheet\d+\.xml)$/;

/**
 * Merge every chart and SmartArt part this slide owns, and the workbooks behind
 * its charts.
 *
 * Returns what it did rather than throwing on what it could not do. A workbook
 * that will not open is a real thing to report — the chart still merged, and
 * the deck is right until somebody edits the data — where a throw would lose
 * the whole run over one unreadable embedding.
 */
export async function mergeGraphics(pkg: Pkg, sites: FieldSite[], resolve: Resolve): Promise<GraphicOutcome> {
  const out = emptyGraphicOutcome();
  // Both lists come from the SITES, which `runPlan` walked once for the whole
  // slide. They used to be two walks of the same relationships — the part list
  // here, and a second lookup walking them again per chart — and the
  // original comment was about keeping both of them ahead of the releases,
  // because releasing between them re-parsed every chart's rels and left them
  // held, one document per record per chart. One walk cannot have that order
  // wrong.
  const graphics = graphicsOf(sites);
  const parts = graphics.map((site) => site.part);
  const workbooks = workbooksOf(sites);

  // BEFORE the workbook's text pass, and that order is load bearing. The
  // numeric pass recognises a value cell by the placeholder still standing in
  // it; `mergeWorkbook` rewrites that shared string to the filled text, after
  // which a cell reading "1250000" is indistinguishable from one somebody typed
  // as a label. Going first is what keeps "the user meant this one to merge"
  // knowable at all.
  for (const site of graphics) {
    // The FIRST package the chart relates to, which is the pairing the numeric
    // pass needs: a value lives in one chart's cache and in one workbook's
    // cell, and filling those from different charts is the mix-up no count
    // would catch.
    if (site.workbooks.length === 0) continue;
    tallyNumbers(out.numbers, await mergeChartNumbers(pkg, site.part, site.workbooks[0], resolve));
  }

  for (const part of parts) out.merged += mergeDocument(await pkg.doc(part), resolve);
  for (const path of workbooks) {
    if (await mergeWorkbook(pkg, path, resolve)) out.workbooks++;
    else out.unreadable.push(path);
  }

  // Written back and dropped, the way `runPlan` drops a finished slide: nothing
  // reads these again, and one live document per chart per record is heap that
  // grows with the row count inside a task-pane WebView.
  for (const part of parts) {
    pkg.release(part);
    pkg.release(Pkg.relsPathFor(part));
  }
  return out;
}

/**
 * Fill the placeholders inside one embedded workbook.
 *
 * A package inside the package: the part is a whole `.xlsx`, so it is opened
 * with its own zip reader, its text parts are merged, and it is written back as
 * bytes. `mergeDocument` is the same reader the slide uses — it knows `<si>` as
 * a text group precisely so a shared string split into `<r><t>` runs merges the
 * way a split paragraph does.
 *
 * Answers false rather than throwing when the bytes are not a readable zip. An
 * embedded workbook can legitimately be something else — an OLE object with the
 * same relationship type, a file another tool wrote — and a merge that loses
 * 240 slides over one of them is worse than a merge that reports it.
 */
async function mergeWorkbook(pkg: Pkg, path: string, resolve: Resolve): Promise<boolean> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(await pkg.bytes(path));
  } catch {
    return false;
  }

  let changed = false;
  for (const name of Object.keys(zip.files)) {
    if (!WORKBOOK_TEXT.test(name)) continue;
    const file = zip.file(name);
    if (!file) continue;
    let doc: Document;
    try {
      doc = parseXml(await file.async("string"));
    } catch {
      // One unparseable part inside an otherwise good workbook. The others are
      // still worth merging, and the chart's own cache already carries the
      // value the reader sees.
      continue;
    }
    if (mergeDocument(doc, resolve) === 0) continue;
    zip.file(name, serializeXml(doc));
    changed = true;
  }
  if (!changed) return true;

  // DEFLATE, because the original is: a workbook written back stored would
  // roughly double the deck's weight per record, on a package the host has to
  // swallow as one base64 string.
  pkg.setBytes(path, await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
  return true;
}

/** Pool one slide's outcome into a run's. */
export function tallyGraphics(into: GraphicOutcome, from: GraphicOutcome): void {
  into.merged += from.merged;
  into.workbooks += from.workbooks;
  tallyNumbers(into.numbers, from.numbers);
  for (const path of from.unreadable) if (!into.unreadable.includes(path)) into.unreadable.push(path);
}
