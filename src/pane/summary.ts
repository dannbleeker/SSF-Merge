/**
 * The sentences the pane puts numbers into.
 *
 * Separate from the step machine because these are the strings a user reads to
 * decide whether to press the one button on the screen, and getting a plural or
 * an off-by-one wrong here is the difference between "720 slides added" and a
 * deck nobody trusts. They are ordinary functions over numbers, so the suite
 * checks the awkward cases — one row, one slide, none — that a hand test with
 * 240 rows never reaches.
 *
 * Every slide reference here is the number the THUMBNAIL RAIL shows. Never an
 * id, which this host refuses for slides a run just added, and never a
 * zero-based index, which is a number the user has no way to see.
 */
import type { Block } from "./steps.js";
import { slidesPerRecord } from "./steps.js";

export function plural(n: number, one: string, many = `${one}s`): string {
  return `${n} ${n === 1 ? one : many}`;
}

/** "Slides 4 to 6" — or "Slide 4" when the block is one slide. */
export function blockName(block: Block): string {
  return block.from === block.to ? `Slide ${block.from}` : `Slides ${block.from} to ${block.to}`;
}

/**
 * What the template block does, in one line.
 *
 * The design's example is "Slides 4 to 6 repeat together, 240 times". "Repeat
 * together" is load-bearing wording: it is what distinguishes a block from
 * three separate templates, and it is the whole reason a record's slides stay
 * adjacent in the output.
 */
export function blockSummary(block: Block, rows?: number): string {
  const name = blockName(block);
  if (rows === undefined) return `${name} — attach data to see how many times this repeats.`;
  const times = rows === 1 ? "once" : `${rows} times`;
  return block.from === block.to ? `${name} repeats ${times}.` : `${name} repeat together, ${times}.`;
}

/** "240 rows × 3 slides" — the arithmetic shown above the merge button. */
export function mergeArithmetic(block: Block, rows: number): string {
  return `${plural(rows, "row")} × ${plural(slidesPerRecord(block), "slide")}`;
}

/** How many slides a merge will add. */
export function slidesAdded(block: Block, rows: number): number {
  return slidesPerRecord(block) * rows;
}

/**
 * What the deck will look like afterwards.
 *
 * Says where they LAND as well as how many, because "720 slides added" answers
 * the wrong half of the question a user hesitating over the button is asking.
 */
export function mergeSummary(block: Block, rows: number, deckSize: number): string {
  const added = slidesAdded(block, rows);
  return `${plural(added, "slide")} added after slide ${deckSize}, leaving ${plural(deckSize + added, "slide")} in the deck.`;
}

/**
 * What an undo will take back.
 *
 * Phrased as the slides themselves rather than as "undo", because the pane is
 * offering to delete part of somebody's presentation and the sentence should
 * say so.
 */
export function undoSummary(added: number, deckSize: number): string {
  if (!undoIsPossible(added, deckSize)) return "Nothing to take back.";
  const from = deckSize - added + 1;
  return added === 1
    ? `Remove slide ${from}, which this merge added.`
    : `Remove slides ${from} to ${deckSize}, which this merge added.`;
}

/**
 * Whether the deck can still contain what the run added.
 *
 * The card is a promise to delete a specific range, and the range is computed
 * backwards from the END of the deck — so a deck SMALLER than the run's own
 * output produces a first slide at or below zero. It read `Remove slides -707
 * to 12, which this merge added.` and offered a button.
 *
 * Reachable, and not only through a bad fixture: the crash crumb offers a run
 * back when the pane reopens, and by then the user may have taken those slides
 * out by hand or with Ctrl+Z. `added` is what the run did; `deckSize` is what
 * is there now; nothing keeps them in step across a closed pane.
 *
 * `sweepPlan` already refuses this case, so pressing the button was safe — it
 * answered "nothing to take back". Safe and wrong: the card said the slides
 * were there and named them. This is the pane agreeing with the decision that
 * will actually be taken.
 */
export function undoIsPossible(added: number, deckSize: number): boolean {
  return added > 0 && deckSize - added + 1 >= 1;
}

/** What a finished merge actually did, as opposed to how much the deck grew. */
export interface MergeReport {
  added: number;
  deckAtStart: number;
  paragraphsMerged?: number;
  skippedRecords?: number;
  skippedSlides?: number;
}

/**
 * The sentence a finished merge gets.
 *
 * "720 slides added" is a measurement of the DECK. It is true of a merge that
 * filled every placeholder and equally true of one that matched none of them
 * and inserted 720 copies of the template — which is the likeliest way a first
 * run against a real template goes wrong, because the placeholders are spelled
 * how the author spelled them and not how the pane expects.
 *
 * So the count of paragraphs actually rewritten goes in the same sentence, and
 * a ZERO is said out loud rather than omitted as an empty clause. `0` is an
 * answer, and it is the whole finding.
 *
 * Clause-joined and zeros otherwise dropped, so an ordinary merge reads as one
 * short sentence and only an unusual one grows.
 */
export function describeMerge(r: MergeReport): string {
  const parts = [`${plural(r.added, "slide")} added after slide ${r.deckAtStart}`];
  if (r.paragraphsMerged !== undefined) {
    parts.push(
      r.paragraphsMerged === 0
        ? "no placeholders were filled — check the spelling in your template"
        : `${plural(r.paragraphsMerged, "placeholder")} filled`,
    );
  }
  // Skips are why "8 rows" and "6 slides" can both be right, and a user who
  // cannot reconcile those two numbers assumes the merge lost something.
  if (r.skippedRecords) parts.push(`${plural(r.skippedRecords, "row")} skipped by a condition`);
  if (r.skippedSlides) parts.push(`${plural(r.skippedSlides, "slide")} skipped by a condition`);
  return `${parts.join(" · ")}.`;
}
