/**
 * Every part a merge touches for one slide, listed once.
 *
 * `prepare` reads these parts to say what placeholders a block holds; `runPlan`
 * and `mergeGraphics` write into them. Each side used to assemble its own list,
 * from its own imports, in its own order — and `prepare.ts` carried the rule
 * they were supposed to obey as a COMMENT:
 *
 *   "The one rule that keeps it from happening a third time is that this list
 *    and `runPlan`'s are the same list."
 *
 * It happened three times. Speaker notes were merged and not scanned, so a
 * block whose placeholders lived in the notes was refused as empty. Then chart
 * labels, the same way. Then a chart's VALUE cells, which live in the workbook
 * the chart relates to — a part the scan did not open at all.
 *
 * Each was fixed by adding the missing part to one of the two lists. This is
 * the fix for the class: there is one list, and a part type that is not in it
 * is invisible to both sides rather than to one.
 */
import type { Pkg } from "../pptx/pkg.js";
import { notesPathFor } from "../pptx/clone.js";
import { graphicPartsOf, packagesOfChart } from "../pptx/graphics.js";

export interface FieldSite {
  /**
   * What kind of part it is, because the two sides do different things with
   * each: only a slide can hold a picture, only a chart has a workbook.
   */
  kind: "slide" | "notes" | "graphic";
  part: string;
  /**
   * For a chart, the packages it relates to — its embedded workbook.
   *
   * A list rather than one, because the two readers want different halves of
   * it and both were separately walking the same relationships to get them.
   * The NUMBERS pass wants the pairing — a value lives in one chart's cache and
   * in one workbook's cell, and filling those from different charts is the
   * mix-up no count would catch — so it takes the first. The workbook TEXT pass
   * wants the slide's whole set, deduplicated, which `workbooksOf` answers.
   *
   * Empty for anything that is not a chart.
   */
  workbooks: string[];
}

/**
 * The sites of one slide, in the order the merge visits them.
 *
 * One walk of the slide's relationships. Before this, `mergeGraphics` walked
 * them for its part list and a second lookup walked them again for the
 * workbooks, which was the same loop twice per chart per record.
 */
export async function fieldSites(pkg: Pkg, slidePath: string): Promise<FieldSite[]> {
  const out: FieldSite[] = [{ kind: "slide", part: slidePath, workbooks: [] }];
  const notes = await notesPathFor(pkg, slidePath);
  if (notes) out.push({ kind: "notes", part: notes, workbooks: [] });
  for (const part of await graphicPartsOf(pkg, slidePath)) {
    out.push({ kind: "graphic", part, workbooks: await packagesOfChart(pkg, part) });
  }
  return out;
}

/** Every workbook the slide's charts relate to, deduplicated, in first-seen order. */
export function workbooksOf(sites: FieldSite[]): string[] {
  const out: string[] = [];
  for (const site of sites) for (const path of site.workbooks) if (!out.includes(path)) out.push(path);
  return out;
}

/** The chart and SmartArt parts, which is what the graphics pass merges. */
export function graphicsOf(sites: FieldSite[]): FieldSite[] {
  return sites.filter((s) => s.kind === "graphic");
}
