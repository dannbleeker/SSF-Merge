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

/** The value the probe writes into the fixture decks' package tags. */
export const PROBE_RUN_TAG = "probe-run";

export interface TagObservation {
  /** What the host said the tag holds, if it found one. */
  value?: string;
  error?: string;
  /**
   * How many slides the insert that was supposed to carry this tag landed.
   *
   * Load-bearing, and the reason this function exists. The probe reads the tag
   * off the LAST slide in the deck, which is the inserted one only if the
   * insert worked. When it did not, that read lands on a slide the user owns,
   * which has never carried our tag and never will.
   */
  insertLanded: number;
}

/**
 * Whether a tag written into the package survives into the host's object model.
 *
 * A never-asked question is not an answer, and reporting one as `no` is worse
 * than reporting nothing: the first sheet this probe took said "the metadata
 * scheme needs rethinking" on the strength of a read that had fallen on the
 * user's own title slide, because the insert in front of it had thrown. The
 * scheme was never tested. That is the same mistake a sibling project's
 * contract gate made with `no-scratch-slide`, and it is guarded here rather
 * than left to the reader to notice.
 */
export function tagVerdict(o: TagObservation): { verdict: Verdict; detail: string } {
  if (!(o.insertLanded > 0)) {
    return {
      verdict: "unknown",
      detail:
        "NOT ASKED — the slide that would carry the tag never landed, so this read fell on a slide the probe did not write. An absent tag here is evidence about nothing.",
    };
  }
  if (o.error !== undefined) return { verdict: "threw", detail: `the read threw: ${o.error}` };
  if (o.value === PROBE_RUN_TAG) {
    return { verdict: "yes", detail: "the whole metadata scheme works, and no tag write is needed in the host" };
  }
  if (o.value === undefined) {
    return {
      verdict: "no",
      detail: "the slide landed and the host did not find the tag, so the metadata scheme needs rethinking",
    };
  }
  return { verdict: "unknown", detail: `the tag is there holding ${JSON.stringify(o.value)}, which nothing wrote` };
}

/**
 * Whose fault a refused insert is.
 *
 * `insertSlidesFromBase64` answering `InvalidArgument` admits two readings that
 * matter in opposite directions: this repo generates a package PowerPoint will
 * not take, or this host takes no package at all. One is a morning's work and
 * the other ends the package path.
 *
 * The control arm is the presentation's OWN bytes, read back through
 * `getFileAsync`. That deck is a package PowerPoint wrote seconds earlier, so
 * it cannot be malformed, and a host that refuses it is refusing insertion
 * itself. This is the partner-question rule: do not reason about which of two
 * readings holds, ask a question only one of them survives.
 */
export function insertionBlame(ours: Verdict, self: Verdict): string {
  if (ours === "yes") return "The insert path works. Whose package it is does not arise.";
  if (self === "yes") {
    return "OURS: this host inserted a deck PowerPoint itself wrote and refused the one we generate. The package writer is at fault; nothing in the failing arms is a fact about the host.";
  }
  if (self === "unknown") {
    return "CANNOT TELL: the control arm never ran, so a refused insert is equally our package or this host. Re-run before concluding anything.";
  }
  return "THE HOST: it refused a deck it wrote itself, so nothing can be inserted here and the package path is blocked for reasons no change in this repo can reach.";
}

/**
 * The two experiments questions three and four run, and what each model
 * predicts. Here rather than in the snippet builder so the reader cannot come
 * to expect a string the probe stopped producing: a guard in
 * `test/probe.test.ts` asserts the generated snippet carries these.
 */
export const Q3 = {
  text: "Hello NAME here",
  /** The one targeted write replaces NAME with Ada. */
  want: "Hello Ada here",
} as const;

export const Q4 = {
  text: "AAA-BBB",
  /**
   * Five characters over AAA, then three over what WAS BBB, both queued in one
   * batch at offsets taken from the ORIGINAL string.
   *
   * If the host evaluates both against the text as it was, the second write
   * lands on BBB. If it evaluates the second against the first one's result,
   * the same offset is three characters to the left. Neither model runs off the
   * end, so a throw here is about something other than the offsets.
   */
  independent: "XXXXX-2",
  shifted: "XXXX2BB",
} as const;

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
