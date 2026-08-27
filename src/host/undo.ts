/**
 * What a run may take back.
 *
 * Undo is POSITIONAL and never by id. A slide a run has just added does not
 * round-trip through `slides.getItem(id)` on PowerPoint for the web, and a
 * sibling project's by-id clean-up once reported 45 successful deletes having
 * removed nothing — leaving 45 blank slides in somebody's presentation and a
 * log that said the job was done.
 *
 * This lived in `verdicts.ts` while only the probe swept. It is the real undo's
 * decision too, and a file about what a probe answer MEANS is the wrong home
 * for the rule that authorises deleting somebody's slides.
 */

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
 * the run started, so nothing the user owned can be reached — and that the
 * deck has gained EXACTLY what the run added, so the last `count` slides are
 * provably the run's own.
 */
export function sweepPlan(o: { deckAtStart: number; deckNow: number; added: number }): SweepPlan | null {
  if (!Number.isInteger(o.deckAtStart) || !Number.isInteger(o.deckNow)) return null;
  if (o.deckAtStart < 0 || o.deckNow < o.deckAtStart) return null;
  const grew = o.deckNow - o.deckAtStart;
  // MORE arrived than this run added, so positional identity is gone.
  //
  // The sweep removes the LAST `count` slides, which is only the run's own
  // slides while nothing else has been appended since. If the deck grew by ten
  // and this run added three, the last three belong to whoever added the other
  // seven — and the run's own three are somewhere in the middle, unreachable
  // by position. Removing anything here deletes a stranger's slides and spares
  // ours.
  //
  // This REVERSES a decision recorded in the tests, and the reversal is the
  // point. The old rule capped the count at `added` on the reasoning that "it
  // is not ours, so do not take it" — sound for a probe that sweeps seconds
  // after it appends, where nothing has had time to arrive. It is not sound
  // for a merge undo, which is a button a user may press after a coffee break
  // on a deck someone else has been editing, and that is the case this
  // function now has to serve.
  //
  // A sibling project's positional sweep met exactly this and refused: its run
  // could account for 68 of the 70 slides the deck had gained, so it left two
  // behind rather than guess. Two blank slides is the right price for never
  // deleting a slide the run cannot prove it created.
  if (grew > o.added) return null;
  const count = Math.min(o.added, grew);
  if (count <= 0) return null;
  const from = o.deckNow - count;
  // Asserted rather than trusted to the arithmetic above, because this call
  // deletes slides from somebody's presentation.
  if (from < o.deckAtStart) return null;
  return { from, count };
}
