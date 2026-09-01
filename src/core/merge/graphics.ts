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

import { SSML_NS, elements, parseXml, serializeXml } from "../pptx/xml.js";
import { mergeDocument, type Resolve } from "./text.js";
import {
  cellAt,
  emptyNumberOutcome,
  mergeChartNumbers,
  tallyNumbers,
  type HeldCell,
  type NumberOutcome,
} from "./numbers.js";
import { graphicsOf, workbooksOf, type FieldSite } from "./sites.js";
import { withinInflatedBudget, workbookParts } from "./workbook.js";

export interface GraphicOutcome {
  /** Placeholders filled in chart and SmartArt parts. */
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
  /** Per workbook, the nodes the numeric pass claimed. */
  const held = new Map<string, HeldCell[]>();
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
    const pass = await mergeChartNumbers(pkg, site.part, site.workbooks[0], resolve);
    tallyNumbers(out.numbers, pass);
    // What that pass REFUSED, kept against the workbook it refused it in. The
    // text pass below reads the same file and would otherwise merge the very
    // placeholder the numeric one declined — see the refusal in `numbers.ts`.
    if (pass.held.length) {
      const path = site.workbooks[0]!;
      held.set(path, [...(held.get(path) ?? []), ...pass.held]);
    }
  }

  for (const part of parts) out.merged += mergeDocument(await pkg.doc(part), resolve);
  for (const path of workbooks) {
    if (await mergeWorkbook(pkg, path, resolve, held.get(path) ?? [])) out.workbooks++;
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
/**
 * Take the nodes the numeric pass claimed out of a workbook part, briefly.
 *
 * A refused value cell keeps the placeholder the author typed — that is the
 * promise — and the text pass that runs next would merge it, in the shared
 * string it reads through or in the cell itself. Neither pass can be reordered:
 * the numeric one has to go first, because it recognises a value cell by the
 * placeholder still standing in it.
 *
 * So the nodes come out for the duration of the merge and go back where they
 * were. Each one is remembered with its parent and the sibling that followed
 * it, which is what makes putting it back exact rather than approximate.
 *
 * **A shared string may serve more than one cell, and holding it holds them
 * all.** Excel keeps one `<si>` per distinct string, so a workbook where the
 * same placeholder is both a value and a label has one entry behind both — and
 * this leaves the label unmerged. That is the right way round: a value cell
 * whose sheet disagrees with the chart is a contradiction nobody sees until
 * they open Edit Data, and it changes the drawing when they close it, where an
 * unmerged label is wrong in a way the author can see. `test/chart-numbers.test.ts`
 * holds the trade so it stays a decision.
 */
function liftHeld(
  doc: Document,
  partName: string,
  isSharedStrings: boolean,
  held: HeldCell[],
): { node: Element; parent: Node; next: Node | null }[] {
  const out: { node: Element; parent: Node; next: Node | null }[] = [];
  const take = (node: Element | undefined) => {
    if (!node?.parentNode) return;
    out.push({ node, parent: node.parentNode, next: node.nextSibling });
    node.parentNode.removeChild(node);
  };
  if (isSharedStrings) {
    const entries = elements(doc, SSML_NS, "si");
    // Highest index first: removing one shifts the positions of those after it,
    // and every index in `held` was read against the untouched table.
    for (const index of held.flatMap((h) => ("si" in h ? [h.si] : [])).sort((a, b) => b - a)) {
      take(entries[index]);
    }
    return out;
  }
  for (const hold of held) {
    if ("si" in hold || hold.sheet !== partName) continue;
    take(cellAt(doc, hold.ref));
  }
  return out;
}

/**
 * The field names a chart's WORKBOOK holds, without filling any of them.
 *
 * `prepareBlock` scanned the chart part and the cells the numeric pass reaches,
 * and nothing else — so a placeholder living only in a workbook cell was filled
 * by the run and unseen by the scan. If it was a block's only placeholder the
 * merge was refused outright, with a sentence telling the author to type field
 * names onto a slide that already carried one: the fourth instance of the class
 * `sites.ts` enumerates, and the direction `prepare.ts` calls the worse one.
 *
 * It bites the generator-written population. A chart PowerPoint authored mirrors
 * its labels into `<c:strCache>`, which the chart-part scan already sees; a
 * workbook whose text is only in the sheet does not.
 *
 * A dry run of `mergeWorkbook` itself, driven by a resolver that records every
 * name and answers null — the same shape `chartValueFields` uses, and for the
 * same reason. Null is what a placeholder with no column gets everywhere else,
 * so nothing merges, `changed` stays false and no workbook is repacked. A
 * second reader that walked the workbook its own way would be free to disagree
 * with the merge about what a placeholder is, which is the disagreement this
 * whole seam exists to prevent.
 */
export async function workbookFields(pkg: Pkg, path: string): Promise<string[]> {
  const seen: string[] = [];
  await mergeWorkbook(
    pkg,
    path,
    (name) => {
      if (!seen.includes(name)) seen.push(name);
      return null;
    },
    [],
  );
  return seen;
}

async function mergeWorkbook(pkg: Pkg, path: string, resolve: Resolve, held: HeldCell[]): Promise<boolean> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(await pkg.bytes(path));
  } catch {
    return false;
  }
  // Refused before a byte is inflated, and reported as unreadable — which is
  // what an unparseable workbook already gets. See `withinInflatedBudget`: this
  // read happens once per merged row, so a bomb costs the row count times its
  // inflated size.
  if (!withinInflatedBudget(zip)) return false;

  // The parts that hold text a merge can fill, taken from what the workbook
  // DECLARES rather than from their names.
  //
  // `sharedStrings.xml` is where Excel keeps every string a cell shows, and it
  // is where a chart's own labels come from in practice. A worksheet is read
  // too because a cell may hold its string INLINE (`<is><t>`), which is what a
  // generator that never built a shared-string table produces.
  //
  // That second half is why the names cannot be matched by pattern, and this
  // used to match `xl/worksheets/sheetN.xml`: a tool that writes its cells
  // differently writes its part names differently, so the population the
  // worksheet read exists for was the population excluded from it. The numeric
  // pass had already learned this and reads the declarations; `workbookParts`
  // is the one reader they now share, so the two cannot disagree about which
  // sheets a workbook has. They did — one filled a value cell while the other
  // left the label cell beside it reading `{{Name}}`.
  //
  // Nothing else is touched. A workbook holds styles, a calc chain and a
  // definition of the range the chart reads; rewriting any of those would
  // change what the chart plots, where this pass only changes what it says.
  const parts = await workbookParts(zip);
  // A workbook whose own declaration will not parse is UNREADABLE, not empty.
  // Both merge nothing, and reporting the second as a success counted a chart
  // among the ones this run had filled while every label in it still read
  // `{{Name}}` in Edit Data.
  if (!parts.readable) return false;
  let changed = false;
  for (const name of [...(parts.sharedStrings ? [parts.sharedStrings] : []), ...parts.sheets]) {
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
    // LIFTED OUT before the merge and put back after, so a node the numeric
    // pass refused is not merged here.
    //
    // Detached rather than skipped, because `mergeDocument` merges a whole
    // document and has no notion of a cell — and detaching is exact where
    // copying the text back would flatten the runs a formatted placeholder is
    // split across. The parent and the next sibling are remembered, so each
    // node returns to the position it left.
    const lifted = liftHeld(doc, name, parts.sharedStrings === name, held);
    const merged = mergeDocument(doc, resolve);
    // LAST OUT, FIRST BACK. Each `next` was the node's live sibling at the
    // moment it was removed, so an anchor is only guaranteed to be in the tree
    // again once every removal after it has been undone. Replaying the list
    // forwards got both halves of this wrong: the shared-string table is taken
    // highest index first, so restoring in that order re-inserted each entry
    // before an anchor the next insertion then jumped in front of and the held
    // strings came back REVERSED — a user's placeholders permuted against the
    // chart's own point order, visible only in Edit Data. And two held cells
    // that are siblings — a row-oriented series with inline strings — made the
    // second removal detach the first's anchor, so `insertBefore` threw
    // "child not in parent", out through `runPlan` and past the pane's own
    // catch: no slides, no notice, an unhandled rejection.
    for (const { node, parent, next } of [...lifted].reverse()) parent.insertBefore(node, next);
    if (merged === 0) continue;
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
