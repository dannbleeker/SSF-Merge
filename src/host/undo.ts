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
 * the run started, so nothing the user owned BEFORE the run can be reached.
 *
 * They do NOT prove the slides are the run's own, and this docstring claimed
 * they did — "the deck has gained EXACTLY what the run added, so the last
 * `count` slides are provably the run's own". Net growth being equal does not
 * mean that. A user who deletes two merged slides and appends two of their own
 * leaves the deck the same size, satisfies every clamp, and gets a plan whose
 * last two entries are slides they made. Every quantity here is a SIZE; none
 * of them is an identity.
 *
 * `provenSweep` is the identity question, asked of the tags in the package.
 * This function answers which slides a sweep may CONSIDER; that one answers
 * which of them it may take.
 */
export function sweepPlan(o: { deckAtStart: number; deckNow: number; added: number }): SweepPlan | null {
  // `added` with the other two, which it was not. The two counts were checked
  // and the third was trusted, and it is the one that decides HOW MANY slides
  // come out: `added: NaN` walked every clamp below untouched — `grew > NaN` is
  // false, `Math.min(NaN, grew)` is NaN, `NaN <= 0` is false, `NaN < deckAtStart`
  // is false — and this returned `{ from: NaN, count: NaN }`. A plan, from the
  // function whose whole job is refusing to produce one.
  //
  // Nothing can reach it today: `added` is a step count and the types carry it.
  // It is checked because the other two are, and because a reader comparing the
  // three has to be able to see one rule rather than guess which is trusted.
  if (!Number.isInteger(o.deckAtStart) || !Number.isInteger(o.deckNow) || !Number.isInteger(o.added)) return null;
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

/**
 * What a SECOND press may ask for, after a first one has run.
 *
 * The pane used to work this out inline, twice, and the two copies answered
 * differently — which is how a fix for a stuck preview became a deletion of a
 * stranger's slides. It is one decision and it is not arithmetic.
 *
 * **A press that DISOWNED anything ends the offer.** `sweepPlan` produces a
 * WINDOW from deck sizes, and carrying `added - removed` widens that window
 * back over the slides the first press just declined. They are then the only
 * slides in it, all untagged — and `provenSweep`'s "a host that answers
 * nothing takes the whole plan" rule, which is right for a host that cannot
 * read tags at all, takes them. Reproduced end to end: preview four slides,
 * delete three by hand, add two of your own, press twice, and the second press
 * deletes the two you added, under a notice saying there was nothing to take
 * back.
 *
 * Narrowing the window instead is what the pane did before, and it is the
 * deadlock: `sweepPlan` refuses a count smaller than the deck's growth, so
 * every later press returns null and the screen that withholds the way on
 * never lets go. Neither number is right, because the question is not how many
 * — it is that a run which has met a slide it cannot claim has lost positional
 * identity for the rest of the range, and no count restores it.
 *
 * **A press that moved NOTHING ends it too.** The same press repeated gives the
 * same answer; offering it again is a button that cannot work, on a screen
 * that may be withholding the way forward.
 *
 * So the offer survives exactly one shape: slides came out, none was declined,
 * and some are still owed. There the window still holds only this run's slides
 * and a second press has less to do.
 */
export function nextSweepOffer(o: { added: number; removed: number; disowned?: number }): number | null {
  if ((o.disowned ?? 0) > 0) return null;
  if (o.removed <= 0) return null;
  const left = o.added - o.removed;
  return left > 0 ? left : null;
}

/**
 * Which of a plan's slides this run may actually delete.
 *
 * `sweepPlan` bounds the RANGE from the deck's sizes. This bounds the SET
 * inside it from what the slides say about themselves: every merged slide
 * carries `SSF_MERGE_RUN` in the package, written before the insert where no
 * host can refuse it, so position stops being the only evidence there is.
 *
 * `tags` is one entry per slide in the plan, in deck order — the run id a slide
 * carries, or `undefined` where it carries none or the host would not say.
 *
 * The answer is 0-based deck indices, HIGHEST FIRST: removing a slide shifts
 * every index after it, so walking upward deletes the wrong slides after the
 * first.
 *
 * Three rules, and each exists because getting it wrong has a cost:
 *
 * - **A host that answers nothing takes the whole plan.** An empty read is not
 *   an empty slide — this repo's rule, learned from two signals agreeing at
 *   zero and both being wrong. A host that cannot read tags is not evidence
 *   that a slide is not ours, so the answer is what it was before anything was
 *   asked. This can only make an undo safer, never less capable.
 * - **A read that does not line up with the plan takes the whole plan.** A
 *   short answer is a read that failed, not a set of slides disowned. Deleting
 *   on it would be acting on an answer nobody gave.
 * - **An id that is not among the answers cannot discriminate**, so the
 *   question falls back to the one that still can: did this add-in make the
 *   slide. A run recovered from a crumb carries a placeholder id, because the
 *   real one was never written — the run died before it could answer. Matching
 *   on it would take back none of the slides that are sitting in the deck,
 *   which is the case recovery exists for.
 */
export function provenSweep(plan: SweepPlan, tags: (string | undefined)[], runId: string): number[] {
  const all: number[] = [];
  for (let i = plan.from + plan.count - 1; i >= plan.from; i--) all.push(i);
  if (tags.length !== plan.count) return all;
  if (tags.every((t) => t === undefined)) return all;
  // Only where the id actually appears can it tell two runs apart.
  const discriminates = tags.includes(runId);
  return all.filter((index) => {
    const tag = tags[index - plan.from];
    if (tag === undefined) return false;
    return discriminates ? tag === runId : true;
  });
}
