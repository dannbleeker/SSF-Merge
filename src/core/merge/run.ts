/**
 * Carry out a plan against a package.
 *
 * Everything happens in the file: clone the template slide, replace the
 * placeholders in the copy, write the merge tags into the copy. Nothing is
 * asked of PowerPoint, which is the point. On the web a slide the run has just
 * added does not resolve by id and a tag write through a shape proxy is
 * refused, so metadata that goes in afterwards is metadata that may not go in
 * at all.
 */
import type { RecordSet } from "../data/recordset.js";
import { cloneSlide, type CloneOptions } from "../pptx/clone.js";
import { cloneSlideGraphics } from "../pptx/graphics.js";
import { Pkg } from "../pptx/pkg.js";
import { writeSlideTags } from "../pptx/tags.js";
import type { MergePlan } from "./plan.js";
import { makeResolver, type EmptyPolicy } from "./resolve.js";
import { mergeDocument } from "./text.js";
import { fieldSites } from "./sites.js";
import { MediaCache, baseName, placeImages, type ImageOutcome, type ResolveImage } from "./images.js";
import { emptyGraphicOutcome, mergeGraphics, tallyGraphics, type GraphicOutcome } from "./graphics.js";

export interface RunOptions {
  onEmpty?: EmptyPolicy;
  /** Passed through to the cloner. Injectable so a test can assert on creation ids. */
  clone?: CloneOptions;
  /**
   * The bytes behind an image field, by FILE NAME.
   *
   * Separate from the records because a record is text and a picture is not.
   * The pane reads the files the user picked and hands them over as a map; the
   * cell holds the name, which is the one thing a spreadsheet can carry.
   *
   * Absent means no image fields can resolve, which is not an error: a template
   * with a picture frame and a merge with no files supplied leaves the
   * placeholder visible and says so.
   */
  images?: Map<string, Uint8Array>;
}

export interface RunResult {
  runId: string;
  /** Package paths of the slides produced, in the order they were added. */
  slides: string[];
  /**
   * How many PLACEHOLDERS were filled by the text passes.
   *
   * A zero on a real template means the fields never matched, which is the
   * alarm the pane speaks from. It counted node GROUPS until 2026-09-01 — one
   * per paragraph rather than one per placeholder — so `Dear {{First}}
   * {{Last}}` reported one; the name is left alone because every caller reads
   * it and the number is what changed.
   *
   * The text passes ONLY. Chart values are counted in `graphics.numbers`, and
   * `describeMerge` withholds the zero alarm when those filled something.
   */
  paragraphsMerged: number;
  /** What the picture pass did, pooled over every slide it touched. */
  images: ImageOutcome;
  /** What the chart and SmartArt pass did, pooled the same way. */
  graphics: GraphicOutcome;
}

/**
 * Run the plan and return what it produced.
 *
 * The order the steps arrive in is the order the slides land in, because every
 * clone is appended to the deck's own slide list. The plan is record-major, so
 * a record's slides stay together.
 */
export async function runPlan(
  pkg: Pkg,
  plan: MergePlan,
  records: RecordSet,
  opts: RunOptions = {},
): Promise<RunResult> {
  const slides: string[] = [];
  let paragraphsMerged = 0;
  const images: ImageOutcome = { placed: 0, missing: [], unreadable: [], stretched: [], crowded: [] };
  const graphics = emptyGraphicOutcome();
  // One cache for the whole run, which is the point of it: a logo on all 240
  // rows becomes ONE media part rather than 240 copies of the same bytes in a
  // package the host has to swallow in a single base64 string.
  const media = new MediaCache(pkg);

  for (const step of plan.steps) {
    const row = records.rows[step.recordIndex];
    if (!row) continue;

    // Clone first, then merge the COPY. Merging the template and cloning after
    // would leave the template holding one record's values, which is the
    // template destroyed rather than used.
    const target = await cloneSlide(pkg, step.source, opts.clone ?? {});
    // The copy's charts and SmartArt, before anything is written into them. A
    // clone inherits the template's relationships wholesale, so without this
    // every record merges into ONE chart part and the whole deck shows the last
    // record's labels — the notes-page defect, a third time.
    await cloneSlideGraphics(pkg, target);
    const resolve = makeResolver(row, { onEmpty: opts.onEmpty });
    // Pictures BEFORE text. To `mergeParagraph` an image field is an ordinary
    // field with a format it does not know, so left to itself it writes the
    // FILE NAME onto the slide; this pass takes the placeholder away and the
    // text pass then finds nothing.
    // One walk of the copy's relationships, shared with the graphics pass and
    // with `prepare`'s scan. The list of parts a merge touches is one list.
    const sites = await fieldSites(pkg, target);
    const slideDoc = await pkg.doc(target);
    tally(images, await placeImages(pkg, target, slideDoc, resolveImage(row, opts.images), media));
    paragraphsMerged += mergeDocument(slideDoc, resolve);
    // The notes page too. It is per-copy content — cloneSlide gives each copy
    // its own precisely so they can differ — and a template whose speaker notes
    // read "Call {{Name}} afterwards" otherwise ships that text verbatim on
    // every merged slide, in the presenter view and on every printed handout.
    const notes = sites.find((site) => site.kind === "notes")?.part;
    if (notes) paragraphsMerged += mergeDocument(await pkg.doc(notes), resolve);
    // Charts and SmartArt. Their text is not on the slide — it is in parts the
    // slide relates to, and in the workbook behind a chart — so `mergeDocument`
    // above never reaches it and this pass is the whole of the feature.
    const drawn = await mergeGraphics(pkg, sites, resolve);
    tallyGraphics(graphics, drawn);
    // Counted into the SAME total the pane reports. `paragraphsMerged` is what
    // becomes "12 placeholders filled", and its zero is the alarm that says a
    // merge added every slide and filled nothing — so a template whose fields
    // all live in a chart would raise that alarm about a merge that worked.
    // `graphics` keeps the breakdown for anyone who needs it.
    paragraphsMerged += drawn.merged;
    await writeSlideTags(pkg, target, step.tags);

    // Nothing reads this slide again, so its parsed copy is written back and
    // dropped. Held, one live document per output slide accumulates on top of
    // the zip's own bytes: 300 clones of a 300-paragraph slide measured 440 MB
    // of heap held against 54 MB released, and 400 clones 591 MB against 54 —
    // flat rather than growing with the record count, which is the property
    // that matters on a task-pane WebView. Every part is byte-identical either
    // way.
    if (notes) {
      pkg.release(notes);
      // Its own .rels too: cloneNotesSlide edits that part to repoint the copy's
      // back-reference at the new slide, so it is parsed and held as well. The
      // guard below caught this one — releasing the slide, its rels and the
      // notes page still left one document per record behind.
      pkg.release(Pkg.relsPathFor(notes));
    }
    pkg.release(Pkg.relsPathFor(target));
    pkg.release(target);
    slides.push(target);
  }

  return { runId: plan.runId, slides, paragraphsMerged, images, graphics };
}

/**
 * A row's image field, as bytes.
 *
 * The cell holds a FILE NAME and the map is keyed by one. Matched
 * case-insensitively and by base name, because `Photos\\ada.PNG` in a
 * spreadsheet and `ada.png` from a file picker are the same picture and a user
 * should not have to know that.
 */
function resolveImage(row: Record<string, string>, images: Map<string, Uint8Array> | undefined): ResolveImage {
  return (name) => {
    if (!images || images.size === 0) return undefined;
    if (!Object.prototype.hasOwnProperty.call(row, name)) return undefined;
    const cell = (row[name] ?? "").trim();
    if (cell === "") return undefined;
    const wanted = baseName(cell);
    for (const [file, bytes] of images) if (baseName(file) === wanted) return bytes;
    return undefined;
  };
}

function tally(into: ImageOutcome, from: ImageOutcome): void {
  into.placed += from.placed;
  for (const key of ["missing", "unreadable", "stretched", "crowded"] as const) {
    for (const name of from[key]) if (!into[key].includes(name)) into[key].push(name);
  }
}
