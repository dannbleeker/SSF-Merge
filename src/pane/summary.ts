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
import type { ImageOutcome } from "../core/merge/images.js";
import type { NumberOutcome } from "../core/merge/numbers.js";
import { sweepPlan } from "../host/undo.js";
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

/**
 * "240 rows × 3 slides" — the arithmetic shown above the merge button.
 *
 * `dropped` is the rows `onEmpty: "skip"` will leave out, and when there are
 * any the heading says BOTH numbers: "228 of 240 rows × 3 slides". Without it
 * the heading multiplies to 720 while the card two lines under it says 684 and
 * the button says 684 — two numbers on one screen, disagreeing, with the
 * smaller one on the thing being pressed. That is exactly the defect the
 * button's own number was fixed for, one element up, and it was found by
 * rendering the screen rather than by any assertion.
 *
 * Both numbers rather than just the survivors: "228 rows" alone is a true
 * sentence that loses how many were pasted, and the gap is the thing a user
 * has to notice.
 */
export function mergeArithmetic(block: Block, rows: number, dropped = 0): string {
  const count = dropped > 0 ? `${rows - dropped} of ${plural(rows, "row")}` : plural(rows, "row");
  return `${count} × ${plural(slidesPerRecord(block), "slide")}`;
}

/**
 * What the deck will look like afterwards.
 *
 * Says where they LAND as well as how many, because "720 slides added" answers
 * the wrong half of the question a user hesitating over the button is asking.
 *
 * FUTURE tense, because nothing has happened yet. It read "6 slides added after
 * slide 3" while sitting directly above an unpressed "Add 6 slides" — the same
 * words the pane uses AFTERWARDS to report what it did, on a screen where it
 * had done nothing at all. A reader taking that sentence at face value has been
 * told their merge already ran, and the only thing telling the two states apart
 * was which screen they happened to be on.
 *
 * The report keeps the past tense and lives in `describeMerge`, so the two no
 * longer read identically at opposite ends of the press.
 */
export function mergeSummary(added: number, deckSize: number | undefined): string {
  // A deck size the host would not give is UNKNOWN, not zero. `?? 0` at the
  // call site read as an empty deck, so on a host whose slide count refuses,
  // the sentence a user reads to decide whether to press said "6 slides will be
  // added after slide 0, leaving 6 slides in the deck" — for a deck with twelve
  // slides in it. Both halves false, in the one sentence that has to be true.
  //
  // The count of slides being added does not depend on the deck's size, so it
  // is still stated; the two clauses that do are dropped rather than invented.
  if (deckSize === undefined) return `${plural(added, "slide")} will be added at the end of the deck.`;
  return `${plural(added, "slide")} will be added after slide ${deckSize}, leaving ${plural(deckSize + added, "slide")} in the deck.`;
}

/**
 * What an undo will take back.
 *
 * Phrased as the slides themselves rather than as "undo", because the pane is
 * offering to delete part of somebody's presentation and the sentence should
 * say so — and the range therefore has to be the range that will actually go.
 *
 * Computed by `sweepPlan`, which is the function the button calls. It used to
 * be computed here, backwards from the end of the deck, and the two disagreed
 * whenever the deck had moved since the merge:
 *
 * - a colleague appends five slides, and the card offered to remove THEIR five
 *   by number while the sweep refused and nothing happened;
 * - the user takes three of the merged slides out by hand, and the card offered
 *   a five-slide range beginning three slides before the merge started;
 * - the user takes all of them out, and the card offered to delete five slides
 *   that pre-date the merge entirely.
 *
 * The button was safe in each — `sweepPlan` refuses what it cannot prove is the
 * run's own — so this was a sentence naming somebody's own slides and a press
 * that did nothing. `main.ts` already said the refusal here was the sweep's, in
 * a comment above a line that did not ask it.
 */
export function undoSummary(added: number, deckSize: number, deckAtStart: number): string {
  const plan = sweepPlan({ deckAtStart, deckNow: deckSize, added });
  if (!plan) return "Nothing to take back.";
  // `sweepPlan` counts from zero, because it is read by `getItemAt`. Every
  // number in this file is the one the thumbnail rail shows.
  const from = plan.from + 1;
  const to = plan.from + plan.count;
  return plan.count === 1
    ? `Remove slide ${from}, which this merge added.`
    : `Remove slides ${from} to ${to}, which this merge added.`;
}

/**
 * Whether there is anything to offer.
 *
 * The same question `sweepPlan` answers, asked of the same function, so the
 * card and the button cannot part company. It began as `deckSize - added + 1 >=
 * 1` — which catches a deck smaller than the run's own output, and is one of
 * three ways the two could disagree.
 */
export function undoIsPossible(added: number, deckSize: number, deckAtStart: number): boolean {
  return sweepPlan({ deckAtStart, deckNow: deckSize, added }) !== null;
}

/** What a finished merge actually did, as opposed to how much the deck grew. */
export interface MergeReport {
  added: number;
  deckAtStart: number;
  paragraphsMerged?: number;
  /**
   * Workbooks behind a chart that could not be opened.
   *
   * Worth a clause of its own because the chart itself is RIGHT: what the
   * reader sees is the merged label, and only "Edit Data" shows the template's
   * placeholder. Silence would leave that to be discovered by a recipient.
   */
  workbooksUnreadable?: number;
  skippedRecords?: number;
  skippedSlides?: number;
  /** Conditions naming a column the data did not have. */
  unknownConditions?: string[];
  /**
   * What became of the pictures, when any were asked for.
   *
   * Same argument as `paragraphsMerged`, on the other half of the merge. The
   * pane's pre-merge tally matches file NAMES and never opens one, so a folder
   * of renamed `.webp` files passes it and places nothing — and the deck grows
   * by 720 slides either way.
   */
  pictures?: ImageOutcome;
  /**
   * What became of the chart values.
   *
   * The sharpest of the three, because the failure is invisible: a value cell
   * whose placeholder did not resolve to a number is left alone, so the chart
   * draws the TEMPLATE's number under the merged label. Nothing about the
   * slide looks wrong.
   */
  chartValues?: NumberOutcome;
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
        ? "no {{fields}} were filled — check the spelling in your template"
        : `${plural(r.paragraphsMerged, "placeholder")} filled`,
    );
  }
  if (r.workbooksUnreadable) {
    parts.push(
      `the data behind ${plural(r.workbooksUnreadable, "chart")} could not be merged — the slides read correctly, ` +
        `but Edit Data still shows your placeholders`,
    );
  }
  // The pictures, when the run had any to place. Silent otherwise, so a
  // text-only merge reads as one short sentence.
  //
  // `missing` is deliberately NOT named. It counts a field whose cell was
  // empty as well as one whose file was never supplied, and an empty cell is
  // documented behaviour — the slide keeps its placeholder, exactly as a text
  // field with no column does. Naming it would report every legitimately blank
  // photo column as a fault. The COUNT carries that case: "2 pictures placed"
  // against 5 rows says it without accusing anybody.
  if (r.pictures) {
    const p = r.pictures;
    const asked = p.placed + p.missing.length + p.unreadable.length + p.crowded.length + p.stretched.length;
    if (asked > 0) {
      // Zero said out loud, for the reason the paragraph count says it: a run
      // that placed no picture at all is the whole finding, and an omitted
      // clause is indistinguishable from a merge with no pictures in it.
      parts.push(p.placed === 0 ? "no pictures were placed" : `${plural(p.placed, "picture")} placed`);
    }
    // The bytes were not an image. Nothing before the merge can catch this:
    // the tally matches names.
    if (p.unreadable.length > 0) {
      parts.push(`the file for ${p.unreadable.join(", ")} is not a picture this add-in can read`);
    }
    // A shape has one fill, so the second field in it was dropped.
    if (p.crowded.length > 0) {
      parts.push(`${p.crowded.join(", ")} had nowhere to go — one shape holds one picture`);
    }
    // Squashed, because the shape inherits its size and there was no ratio to
    // fit to. Reads as a broken image and is not one.
    if (p.stretched.length > 0) {
      parts.push(`${p.stretched.join(", ")} was stretched to fit a shape that states no size`);
    }
  }
  // The chart values, when the run had any to fill.
  if (r.chartValues) {
    const c = r.chartValues;
    // No zero clause here, and the asymmetry with the pictures above is the
    // point. A value placeholder that was LOOKED AT either fills or refuses, so
    // a zero between those two is always a refusal — and the refusal clause
    // below says the same thing and names the count. "no chart values were
    // filled · 3 chart values did not read as a number" is one fact reported
    // twice. A picture can be absent for a reason this sentence deliberately
    // does not name (an empty cell keeps its placeholder by design), so there
    // the zero is the only thing said and has to stay.
    //
    // "Looked at" is the correction, and it used to read "every". A third
    // outcome exists: a series whose range cannot be read, or which names a
    // sheet the workbook does not declare, is abandoned before any cell is
    // opened. That was a bare `continue` in `numbers.ts` — no fill, no refusal,
    // nothing — so this block was handed two zeros and said nothing at all,
    // about the one case where the user has no other signal. It is counted now
    // and it has its own clause.
    if (c.filled > 0) parts.push(`${plural(c.filled, "chart value")} filled`);
    // The one a reader cannot see for themselves. Nothing is written when a
    // cell's placeholder does not resolve to a number, so the point keeps the
    // template's — a chart that is wrong under a label that is right.
    if (c.refused > 0) {
      parts.push(
        `${plural(c.refused, "chart value")} did not read as a number, so ${
          c.refused === 1 ? "that point still shows" : "those points still show"
        } the template's`,
      );
    }
    // Said in the language of the CHART rather than of the cell, because when a
    // range cannot be read there is no cell to point the reader at — the count
    // is per series, which is the only honest granularity for it.
    // The one with a REMEDY, so the sentence carries it. The value is in the
    // workbook — an ordinary text merge put it there — and the chart's cached
    // point list has no entry at that index to draw it from, because the writer
    // omitted the point for a cell that held no number. Opening Edit Data and
    // closing it refreshes the cache from the sheet, which is a thing the user
    // can do and the merge cannot.
    if (c.unplotted > 0) {
      parts.push(
        `${plural(c.unplotted, "chart value")} reached the data sheet but not the chart — press Edit Data on ${
          c.unplotted === 1 ? "that chart" : "those charts"
        } and close it to bring ${c.unplotted === 1 ? "it" : "them"} in`,
      );
    }
    if (c.unreadable > 0) {
      parts.push(
        `the data behind ${plural(c.unreadable, "chart series")} could not be read, so ${
          c.unreadable === 1 ? "it keeps" : "they keep"
        } the template's numbers`,
      );
    }
  }
  // Skips are why "8 rows" and "6 slides" can both be right, and a user who
  // cannot reconcile those two numbers assumes the merge lost something.
  // Two different reasons, and they said the same one. A SLIDE is left out by a
  // condition; a ROW is left out only under `onEmpty: "skip"`, which drops a
  // record when a field on one of its slides has no value — `buildPlan` fills
  // `skippedRecords` from nowhere else. Telling a user their row was dropped
  // by a condition sends them to the condition control, where there is nothing
  // to find.
  if (r.skippedRecords) parts.push(`${plural(r.skippedRecords, "row")} skipped for a blank field`);
  if (r.skippedSlides) parts.push(`${plural(r.skippedSlides, "slide")} skipped by a condition`);
  // The condition that did NOTHING, which is the one worth saying. The engine
  // emits the slide anyway rather than hiding an authoring mistake behind
  // output that looks finished — so without this line the user sees the right
  // number of slides and never learns the condition was ignored.
  if (r.unknownConditions && r.unknownConditions.length > 0) {
    parts.push(`no column for ${r.unknownConditions.join(", ")}, so those slides were included for every row`);
  }
  return `${parts.join(" · ")}.`;
}
