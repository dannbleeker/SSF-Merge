/**
 * Turn what the pane knows into what the engine needs, and say no clearly.
 *
 * The pane speaks in SLIDE NUMBERS — the ones in the thumbnail rail, which is
 * the only numbering a user can see. The engine speaks in package paths. This
 * is the translation, and it is pure so that every refusal can be checked
 * without a PowerPoint: the alternative is a merge that starts, clones the
 * wrong slides, and is discovered by looking at the output.
 */
import type { Pkg } from "../pptx/pkg.js";
import type { Block, BlockSlide } from "./plan.js";
import { notesPathFor } from "../pptx/clone.js";
import { graphicPartsOf } from "../pptx/graphics.js";
import { fieldsIn } from "./text.js";
import { chartValueFields } from "./numbers.js";
import { imageFieldsIn } from "./images.js";
import { workbookOfChart } from "../pptx/graphics.js";

export interface BlockRequest {
  /** First slide of the template, 1-based, as the thumbnail rail shows it. */
  from: number;
  /** Last slide, 1-based, inclusive. */
  to: number;
  /** Where the block starts INSIDE the package that was read back. */
  offsetInPackage: number;
  /** Conditions the pane collected, keyed by slide number. */
  conditions?: Record<number, string>;
  /**
   * Whether a block with no `{{fields}}` on it is an ANSWER rather than a refusal.
   *
   * A merge with no placeholders produces N identical copies, so `runMerge`
   * must never accept one — and it does not pass this. But the pane now picks
   * the slides BEFORE the fields are on them: the order is choose slides,
   * paste data, insert fields, merge, because the names to type are the data's
   * own column headers and the user cannot know them at step 1. That was
   * reported from a first real run: the refusal told somebody to go and type
   * field names nobody had yet.
   *
   * So the same read serves two questions. "What is on these slides" tolerates
   * nothing being on them; "may I clone them 240 times" does not. Passed
   * explicitly rather than inferred from `runId === "inspect"`, because the one
   * thing this flag must never do is turn off by accident on the destructive
   * path.
   */
  allowEmpty?: boolean;
}

export type Prepared =
  | {
      ok: true;
      block: Block;
      fields: string[];
      /**
       * The fields written as a PICTURE — `{{Photo|image}}` and its two
       * siblings — by name.
       *
       * Separate from `fields` because the pane needs a different answer from
       * it. `imageFieldsIn` has answered this question since it was written and
       * nothing called it, so the pane decided which columns were pictures from
       * the DATA's detected types alone, while the engine decides from the
       * FIELD's format. See `imagesWanted`.
       */
      imageFields: string[];
    }
  | { ok: false; why: string };

/**
 * Build the block, or refuse with a sentence the pane can show as it stands.
 *
 * Every refusal here is a thing the user can act on. "Something went wrong" is
 * the failure this function exists to avoid: a merge is expensive to undo and
 * the user is the only one who knows which slides they meant.
 */
export async function prepareBlock(pkg: Pkg, req: BlockRequest, runId: string): Promise<Prepared> {
  if (!Number.isInteger(req.from) || !Number.isInteger(req.to)) {
    return { ok: false, why: "The template block has to be whole slide numbers." };
  }
  if (req.from < 1) return { ok: false, why: "Slides are numbered from 1." };
  if (req.to < req.from) {
    return { ok: false, why: `The block ends before it starts: slide ${req.from} to ${req.to}.` };
  }

  const paths = await pkg.slidePaths();
  // The package read back is either the whole deck or just the block, and the
  // two number their slides differently. `offsetInPackage` is what reconciles
  // them; getting it wrong merges the wrong slides and the output looks
  // deliberate, which is why `templateOffset` is its own tested function.
  const start = req.offsetInPackage;
  const count = req.to - req.from + 1;
  if (start < 0 || start + count > paths.length) {
    return {
      ok: false,
      why: `The template block is slides ${req.from} to ${req.to}, and the deck that came back has ${paths.length}.`,
    };
  }

  const slides: BlockSlide[] = [];
  const fields: string[] = [];
  const imageFields: string[] = [];
  for (let i = 0; i < count; i++) {
    const path = paths[start + i];
    if (!path) return { ok: false, why: `Slide ${req.from + i} is not in the deck that came back.` };
    // The slide, and its speaker notes.
    //
    // `runPlan` merges the notes page — a template whose notes read "Call
    // {{Name}} afterwards" otherwise ships that verbatim on every handout — but
    // this scan only ever read the slide. So a block whose placeholders live in
    // the notes was refused with "no placeholders, so every copy would be
    // identical", about a merge that would have filled them.
    //
    // The mirror of the chart case below, and the worse direction: there the
    // pane reported fields it cannot merge, here it hid fields it can and
    // blocked the merge on the strength of it.
    const notes = await notesPathFor(pkg, path);
    // The slide, its speaker notes, and the charts and SmartArt it shows.
    //
    // All three are places `runPlan` fills, so all three belong in the same
    // list: a scan that reads fewer parts than the merge writes REFUSES a block
    // it would have merged, with "no placeholders" about a slide the author is
    // looking at a placeholder on. That happened twice — first for notes, then
    // for charts, which were reported as unfillable right up until they became
    // fillable. The one rule that keeps it from happening a third time is that
    // this list and `runPlan`'s are the same list.
    //
    // A chart's WORKBOOK is read for one thing only: the VALUE cells.
    //
    // It used to be skipped entirely, on the reasoning that its strings are the
    // chart's own cache strings and would name nothing new. That was true until
    // chart numbers became a feature. A value cell holds its placeholder in the
    // WORKBOOK — `fieldsIn` cannot see it, because `<c:numCache>` is excluded
    // from the text pass on purpose, a formatted number being unplottable.
    //
    // So a slide whose only field was the number its chart plots was refused
    // with "has no {{fields}}, so every copy would be identical", about a merge
    // that filled it. The third time this scan has read fewer parts than the
    // merge writes, after notes and after chart labels, and the first two are
    // described a few lines above.
    //
    // The labels still come from the cache, not the workbook: those two hold
    // the same strings and reading both would name nothing twice.
    const slideDoc = await pkg.doc(path);
    for (const name of imageFieldsIn(slideDoc)) if (!imageFields.includes(name)) imageFields.push(name);
    const own = [...fieldsIn(slideDoc), ...(notes ? fieldsIn(await pkg.doc(notes)) : [])];
    for (const part of await graphicPartsOf(pkg, path)) {
      own.push(...fieldsIn(await pkg.doc(part)));
      if (!part.startsWith("ppt/charts/")) continue;
      own.push(...(await chartValueFields(pkg, part, await workbookOfChart(pkg, part))));
    }
    for (const f of own) if (!fields.includes(f)) fields.push(f);
    const condition = req.conditions?.[req.from + i];
    slides.push({ path, seq: i + 1, fields: own, ...(condition ? { condition } : {}) });
  }

  if (fields.length === 0 && !req.allowEmpty) {
    // Not an error the engine can see: a merge with no placeholders produces N
    // identical copies, which is never what anybody meant and is expensive to
    // undo once it is in the deck.
    const where = count === 1 ? `Slide ${req.from}` : `Slides ${req.from} to ${req.to}`;
    // Says the SYNTAX, not the word "placeholder".
    //
    // PowerPoint calls its own empty content boxes placeholders — "Click to add
    // title" IS a placeholder in its vocabulary — so a user looking at two of
    // them was being told there are none. First contact with this add-in, on a
    // fresh deck, is exactly the moment that reads as the thing being broken.
    return {
      ok: false,
      why:
        `${where} has no {{fields}}, so every copy would be identical. Type your column headers onto ` +
        `the slides in double braces — {{First}}, {{City}} — then press again. PowerPoint's own empty ` +
        `"Click to add title" boxes are not fields.`,
    };
  }

  return { ok: true, block: { id: runId, slides }, fields, imageFields };
}
