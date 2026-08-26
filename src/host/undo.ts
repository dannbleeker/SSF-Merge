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
