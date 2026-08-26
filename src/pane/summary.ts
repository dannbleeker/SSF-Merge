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
  if (added <= 0) return "Nothing to take back.";
  const from = deckSize - added + 1;
  return added === 1
    ? `Remove slide ${from}, which this merge added.`
    : `Remove slides ${from} to ${deckSize}, which this merge added.`;
}
