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
import { fieldsIn } from "./text.js";
import { chartValueFields } from "./numbers.js";
import { workbookFields } from "./graphics.js";
import JSZip from "jszip";

import { imageFieldsIn } from "./images.js";
import { fieldSites } from "./sites.js";

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
      /** Picture fields written where no picture can be placed — a notes page, a chart, SmartArt. */
      imageFieldsOffSlide: string[];
      /**
       * The fields on each slide of the block, in `seq` order.
       *
       * `fields` above is the block's whole set, flattened, and that is the
       * right answer for the pane's chip list. It is the WRONG answer for
       * `onEmpty: "skip"`, which drops a record when a field on one of the
       * slides this record actually gets is blank — so a pane holding only the
       * flat list would skip a row over a blank field on a slide that row's own
       * condition had already removed, and put a number on the merge button
       * the plan does not produce.
       *
       * Already computed per slide (`own`); it was thrown away at the join.
       */
      slideFields: string[][];
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
    // "the deck that came back" is only true on the whole-file route. On the
    // subset route PowerPoint sends back the exported BLOCK, so the number here
    // is the size of that export and matches nothing the user can count in
    // their own deck — a sentence naming a deck of 4 to somebody looking at 30
    // slides.
    //
    // `templateOffset` is `0` for the subset route and the block's start for the
    // file route, so a positive offset PROVES the whole deck came back. Zero is
    // either — the subset route, or the file route on a block starting at slide
    // 1 — and the neutral phrase is true of both, which is the direction to be
    // wrong in.
    const what = start > 0 ? "the deck that came back" : "the slides PowerPoint sent back";
    return {
      ok: false,
      why: `The template block is slides ${req.from} to ${req.to}, and ${what} has ${paths.length}.`,
    };
  }

  const slides: BlockSlide[] = [];
  const fields: string[] = [];
  const imageFields: string[] = [];
  // Picture fields written somewhere `placeImages` cannot fill: a notes page, a
  // chart's text, a SmartArt node. They are collected so they can be REPORTED —
  // see `imageFieldsOffSlide`.
  const imageFieldsOffSlide: string[] = [];
  for (let i = 0; i < count; i++) {
    const path = paths[start + i];
    if (!path)
      return {
        ok: false,
        why: `Slide ${req.from + i} is not in ${start > 0 ? "the deck that came back" : "the slides PowerPoint sent back"}.`,
      };
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
    // ONE list, the same one `runPlan` merges into. Each of the three times
    // this scan and that merge disagreed, the fix was to add a part to one of
    // two hand-assembled lists; `fieldSites` is the fix for the class, and a
    // part type missing from it is invisible to both sides rather than to one.
    const own: string[] = [];
    for (const site of await fieldSites(pkg, path)) {
      const doc = await pkg.doc(site.part);
      own.push(...fieldsIn(doc));
      // Only a slide can hold a picture — `placeImages` fills a SHAPE, and a
      // chart part has none.
      if (site.kind === "slide") {
        for (const name of imageFieldsIn(doc)) if (!imageFields.includes(name)) imageFields.push(name);
      } else {
        // NAMED rather than ignored. `{{Photo|image}}` on a notes page merges as
        // nothing and prints itself: the raw placeholder reaches presenter view
        // and every handout, and if it is the block's only picture field the
        // pane never even offers the file picker. Filling it is not on the
        // table — `placeImages` fills a shape and a notes page is not one — so
        // the answer is to say so before the merge rather than after it.
        for (const name of imageFieldsIn(doc)) if (!imageFieldsOffSlide.includes(name)) imageFieldsOffSlide.push(name);
      }
      // A chart's VALUE cells live in the workbook it relates to, and the
      // reader is a dry run of the merge's own walk, so the two cannot hold
      // different opinions about which cells carry a placeholder.
      //
      // ONE inflate per workbook, shared by both readers below. They read
      // different things — the value cells and the text — out of the same zip,
      // and inflating it twice doubled the cost of this step on any deck with
      // charts. Safe only because both are DRY RUNS whose resolver writes
      // nothing: a shared book on the merge path would let one pass see the
      // other's edits. See `mergeChartNumbers`.
      const inflated = new Map<string, JSZip>();
      for (const book of site.workbooks) {
        if (!pkg.has(book)) continue;
        try {
          inflated.set(book, await JSZip.loadAsync(await pkg.bytes(book)));
        } catch {
          // Unreadable here is unreadable in the readers too, and they already
          // answer for it. Left out of the map so they take their own path.
        }
      }
      if (site.workbooks.length > 0)
        own.push(...(await chartValueFields(pkg, site.part, site.workbooks[0], inflated.get(site.workbooks[0] ?? ""))));
      // And the workbook's own TEXT, which the numeric walk above does not
      // reach: it opens only the cells a `<c:f>` names and the cache has a
      // point for. A placeholder in any other cell was filled by the run and
      // unseen here — and if it was the block's only one, the merge was refused
      // as though the slide had no fields at all.
      //
      // Every workbook, not `[0]`: the text pass merges the whole set.
      for (const book of site.workbooks) own.push(...(await workbookFields(pkg, book, inflated.get(book))));
    }
    // Deduped here rather than at each push. A chart's label is reported by the
    // chart-part scan AND by its workbook, which is not a disagreement — it is
    // the same string in the two places PowerPoint keeps it.
    for (const f of own) if (!fields.includes(f)) fields.push(f);
    const condition = req.conditions?.[req.from + i];
    slides.push({ path, seq: i + 1, fields: own, ...(condition ? { condition } : {}) });
  }

  if (fields.length === 0 && !req.allowEmpty) {
    // Not an error the engine can see: a merge with no placeholders produces N
    // identical copies, which is never what anybody meant and is expensive to
    // undo once it is in the deck.
    // The verb with the subject, because this is the one place that knows
    // whether the subject is singular. It said "Slides 2 to 4 HAS no
    // {{fields}}" — the mirror of the pane's own "Slide 4 carry no fields yet",
    // in the engine refusal the pane shows verbatim. `blockCarries` in the pane
    // already solved this; this sentence never got the fix.
    const where = count === 1 ? `Slide ${req.from} has` : `Slides ${req.from} to ${req.to} have`;
    // Says the SYNTAX, not the word "placeholder".
    //
    // PowerPoint calls its own empty content boxes placeholders — "Click to add
    // title" IS a placeholder in its vocabulary — so a user looking at two of
    // them was being told there are none. First contact with this add-in, on a
    // fresh deck, is exactly the moment that reads as the thing being broken.
    return {
      ok: false,
      why:
        `${where} no {{fields}}, so every copy would be identical. Type your column headers onto ` +
        `the slides in double braces — {{First}}, {{City}} — then press again. PowerPoint's own empty ` +
        `"Click to add title" boxes are not fields.`,
    };
  }

  return {
    ok: true,
    block: { id: runId, slides },
    fields,
    imageFields,
    imageFieldsOffSlide,
    slideFields: slides.map((s) => s.fields ?? []),
  };
}
