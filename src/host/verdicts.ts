/**
 * What an observation means, decided away from the host.
 *
 * A probe that reasons inside a `PowerPoint.run` callback is a probe whose
 * conclusions cannot be tested. Every judgement this one makes lives here, as a
 * pure function over numbers and strings, so the rules can go red in CI while
 * the round trip through a real PowerPoint stays a matter of reading values
 * out.
 *
 * The wording matters as much as the logic. A sibling project spent two
 * sessions reasoning about answers that could equally have been about the host
 * or about the thing asking, so every verdict here either names a mechanism or
 * says plainly that it cannot tell.
 */

export type Verdict = "yes" | "no" | "unknown" | "threw";

export interface InsertObservation {
  /** Deck size before the call, measured. */
  before: number;
  /** Deck size after the call, measured in its own context. */
  after: number;
  /** How many slides the file being inserted actually contains. */
  expected: number;
  /** The error, if the call threw. */
  error?: string;
}

export interface InsertVerdict {
  verdict: Verdict;
  landed: number;
  detail: string;
}

/**
 * Whether an insert did what it said.
 *
 * The delta is the evidence, never the absence of an error: this host family
 * has accepted slide adds, deck inserts and tag writes and performed none of
 * them. A partial landing is called out separately from nothing landing,
 * because they point at different things.
 */
export function insertVerdict(o: InsertObservation): InsertVerdict {
  if (o.error !== undefined) {
    return { verdict: "threw", landed: 0, detail: `the call threw: ${o.error}` };
  }
  const landed = o.after - o.before;
  if (landed === o.expected) {
    return { verdict: "yes", landed, detail: `all ${o.expected} slide(s) landed` };
  }
  if (landed === 0) {
    return {
      verdict: "no",
      landed,
      detail: `the call raised nothing and the deck did not grow, so ${o.expected} slide(s) were dropped silently`,
    };
  }
  return {
    verdict: "no",
    landed,
    detail: `${landed} of ${o.expected} slide(s) landed, which is a partial insert rather than a refusal`,
  };
}

/**
 * The headline for question one, from both arms.
 *
 * Asking only the fresh-id arm cannot separate "office-js#6105 does not
 * reproduce here" from "this host refuses every insert", so the collision arm
 * runs in the same session and the reading is taken from the pair. That is the
 * partner-question rule: a probe answer that admits two readings is not an
 * answer.
 */
export function creationIdReading(fresh: InsertVerdict, collision: InsertVerdict): string {
  if (fresh.verdict === "yes" && collision.verdict !== "yes") {
    return "CONFIRMED: a fresh creation id inserts and a duplicated one does not. The engine's rewrite is what makes cloning work, and office-js#6105 reproduces here.";
  }
  if (fresh.verdict === "yes" && collision.verdict === "yes") {
    return "Both arms landed. The creation id rewrite is harmless but this host does not need it: office-js#6105 does not reproduce here. Keep the rewrite, since the hosts that do need it are the ones we cannot test.";
  }
  if (fresh.verdict !== "yes" && collision.verdict !== "yes") {
    return "BLOCKING: neither arm landed, so this says nothing about creation ids and everything about insertion. Read the errors before building on the package path.";
  }
  return "Unexpected: the duplicated id landed and the fresh one did not. Do not build on either reading; re-run before drawing any conclusion.";
}

export interface SubstringObservation {
  /** Text of the shape before the write. */
  before: string;
  /** Text after the write. */
  after: string;
  /** What the write was supposed to produce. */
  want: string;
  /** Whether the styled span was still styled afterwards. */
  boldAfter?: boolean;
  error?: string;
}

/**
 * Whether a substring write kept the formatting around it.
 *
 * Two things are asked at once on purpose, because a write that lands in the
 * wrong place and a write that flattens the run look the same in a screenshot.
 */
export function substringVerdict(o: SubstringObservation): { verdict: Verdict; detail: string } {
  if (o.error !== undefined) return { verdict: "threw", detail: `the write threw: ${o.error}` };
  if (o.after !== o.want) {
    return {
      verdict: "no",
      detail: `the text came out as ${JSON.stringify(o.after)} rather than ${JSON.stringify(o.want)}, so the offsets are not what the caller thinks`,
    };
  }
  if (o.boldAfter === undefined)
    return {
      verdict: "unknown",
      detail: "the text is right; the host would not say whether the run kept its formatting",
    };
  return o.boldAfter
    ? { verdict: "yes", detail: "the text is right and the styled run is still styled, so a targeted write is safe" }
    : {
        verdict: "no",
        detail: "the text is right but the styling is gone, so a targeted write flattens the run it lands in",
      };
}

/**
 * Whether two writes queued in one batch interfere.
 *
 * The answer decides whether replacements must be queued right to left or can
 * be issued in any order, which is the difference between a correct in-place
 * preview and one that silently scrambles a sentence.
 */
export function offsetVerdict(
  after: string,
  ifIndependent: string,
  ifShifted: string,
): { verdict: Verdict; detail: string } {
  if (after === ifIndependent) {
    return {
      verdict: "yes",
      detail: "both writes used the ORIGINAL offsets, so replacements may be queued in any order",
    };
  }
  if (after === ifShifted) {
    return {
      verdict: "no",
      detail: "the second write saw the first one's result, so replacements must be queued right to left",
    };
  }
  return {
    verdict: "unknown",
    detail: `neither model predicts ${JSON.stringify(after)}. Record it and do not guess: queue right to left, which is correct under both.`,
  };
}

export interface SweepPlan {
  from: number;
  count: number;
}

/**
 * Which slides the probe may remove when it is done.
 *
 * Positional, never by id: a slide this run added does not round-trip through
 * `slides.getItem(id)` on the web, and a sibling project's clean-up reported 45
 * successful deletes having removed nothing, leaving 45 blank slides in
 * somebody's deck.
 *
 * Every clamp here is load-bearing and each is asserted in the tests. Together
 * they guarantee the first index removed is at or after the deck's size when
 * the probe started, so nothing the user owned can be reached.
 */
export function sweepPlan(o: { deckAtStart: number; deckNow: number; added: number }): SweepPlan | null {
  if (!Number.isInteger(o.deckAtStart) || !Number.isInteger(o.deckNow)) return null;
  if (o.deckAtStart < 0 || o.deckNow < o.deckAtStart) return null;
  const grew = o.deckNow - o.deckAtStart;
  const count = Math.min(o.added, grew);
  if (count <= 0) return null;
  const from = o.deckNow - count;
  // Asserted rather than trusted to the arithmetic above, because this call
  // deletes slides from somebody's presentation.
  if (from < o.deckAtStart) return null;
  return { from, count };
}
