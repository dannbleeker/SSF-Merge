import { describe, expect, it } from "vitest";
import { provenSweep, sweepPlan } from "../src/host/undo.js";

describe("sweepPlan", () => {
  it("removes only what this run added", () => {
    expect(sweepPlan({ deckAtStart: 10, deckNow: 14, added: 4 })).toEqual({ from: 10, count: 4 });
  });

  it("never removes more than the deck actually grew", () => {
    // The probe thinks it added six; the deck only grew by two. Believing the
    // probe would reach four slides back into the user's own content.
    expect(sweepPlan({ deckAtStart: 10, deckNow: 12, added: 6 })).toEqual({ from: 10, count: 2 });
  });

  it("refuses outright when MORE arrived than the run added", () => {
    // Reversed on 2026-08-27, when undo became a button a user can press
    // rather than a probe's own clean-up.
    //
    // This asserted `{ from: 17, count: 3 }` — take the last three — on the
    // reasoning that the extra slides "are not ours, so do not take them".
    // That reasoning inverts: the sweep removes the LAST slides, so if the
    // deck grew by ten and this run added three, the last three belong to
    // whoever added the other seven. It would delete a stranger's slides and
    // leave ours in place.
    //
    // Harmless for a probe that sweeps seconds after it appends. Not harmless
    // for an undo pressed after a coffee break on a deck a co-author has been
    // editing, which is what this now has to serve.
    expect(sweepPlan({ deckAtStart: 10, deckNow: 20, added: 3 })).toBeNull();
  });

  it("still sweeps when the deck gained exactly what the run added", () => {
    // The ordinary case, and the one the refusal must not break.
    expect(sweepPlan({ deckAtStart: 10, deckNow: 13, added: 3 })).toEqual({ from: 10, count: 3 });
  });

  it("still sweeps when the host took FEWER than the run asked for", () => {
    // A short insert is the run's own shortfall, not a stranger's arrival, and
    // the slides that did land are still the last ones.
    expect(sweepPlan({ deckAtStart: 10, deckNow: 12, added: 6 })).toEqual({ from: 10, count: 2 });
  });

  it("does nothing when the deck did not grow", () => {
    expect(sweepPlan({ deckAtStart: 10, deckNow: 10, added: 4 })).toBeNull();
  });

  it("does nothing when the deck SHRANK, rather than computing a negative index", () => {
    expect(sweepPlan({ deckAtStart: 10, deckNow: 8, added: 4 })).toBeNull();
  });

  it("refuses a plan built from counts that are not counts", () => {
    expect(sweepPlan({ deckAtStart: Number.NaN, deckNow: 12, added: 2 })).toBeNull();
  });

  it("never plans a first index before the deck's size at the start", () => {
    // The invariant the whole function exists for, asserted directly over a
    // range rather than inferred from the arithmetic.
    for (let start = 0; start < 30; start++) {
      for (let now = start; now < start + 30; now++) {
        for (let added = 0; added < 30; added++) {
          const plan = sweepPlan({ deckAtStart: start, deckNow: now, added });
          if (plan) {
            expect(plan.from).toBeGreaterThanOrEqual(start);
            expect(plan.from + plan.count).toBeLessThanOrEqual(now);
          }
        }
      }
    }
  });
});

describe("what sweepPlan may never produce", () => {
  /**
   * This function authorises deleting slides from somebody's presentation, so
   * the properties are asserted over the whole space rather than at the points
   * a reader thought of.
   *
   * Four rules, and each is the reason a specific clamp is in there:
   * a plan may not reach a slide the user owned before the run, may not reach
   * past the end of the deck, may not remove more than the run added, and may
   * not be empty or fractional. `getItemAt` takes an index; a fractional one is
   * not a slide.
   */
  it("holds over every combination of the three counts", () => {
    const violations: string[] = [];
    for (let deckAtStart = 0; deckAtStart <= 12; deckAtStart++) {
      for (let deckNow = 0; deckNow <= 20; deckNow++) {
        for (let added = 0; added <= 12; added++) {
          const plan = sweepPlan({ deckAtStart, deckNow, added });
          // The rule that makes the last `count` slides provably this run's
          // own. Stated as a property because the four below cannot see it: a
          // plan built after a stranger appended still starts past
          // `deckAtStart` and still ends at the deck's end, and is still
          // somebody else's slides.
          if (deckNow - deckAtStart > added && plan) {
            violations.push(`${deckAtStart}/${deckNow}/${added}: swept a deck that grew by more than the run added`);
            continue;
          }
          if (!plan) continue;
          const bad: string[] = [];
          if (plan.from < deckAtStart) bad.push("reaches a slide the user owned");
          if (plan.from + plan.count > deckNow) bad.push("reaches past the end of the deck");
          if (plan.count > added) bad.push("removes more than the run added");
          if (plan.count <= 0) bad.push("an empty plan returned as a plan");
          if (!Number.isInteger(plan.from) || !Number.isInteger(plan.count)) bad.push("a fractional index");
          if (bad.length) violations.push(`${deckAtStart}/${deckNow}/${added}: ${bad.join(", ")}`);
        }
      }
    }
    expect(violations).toEqual([]);
  });

  it("refuses a count that is not a whole number of slides", () => {
    /**
     * The two deck counts were checked and `added` was trusted, and `added` is
     * the one that decides how many slides come out. `NaN` walked every clamp
     * untouched — `grew > NaN` is false, `Math.min(NaN, grew)` is NaN,
     * `NaN <= 0` is false, `NaN < deckAtStart` is false — and a plan of
     * `{ from: NaN, count: NaN }` came back out of the function whose whole job
     * is refusing to produce one.
     *
     * Unreachable through the types today. Checked because the other two are.
     */
    for (const added of [Number.NaN, 2.5, Number.POSITIVE_INFINITY, -1]) {
      expect(sweepPlan({ deckAtStart: 5, deckNow: 8, added }), String(added)).toBeNull();
    }
    // And the whole-number case still works, so this refuses shapes not counts.
    expect(sweepPlan({ deckAtStart: 5, deckNow: 8, added: 3 })).toEqual({ from: 5, count: 3 });
  });
});

describe("proving a slide is this run's before deleting it", () => {
  /**
   * `sweepPlan`'s clamps compare SIZES, never identity, and its own docstring
   * claimed more than they deliver: "the deck has gained EXACTLY what the run
   * added, so the last `count` slides are provably the run's own". Net growth
   * being equal does not mean that.
   *
   * Deck of 12. A merge adds 6, making 18. The user reviews the output — which
   * is the whole point of a mail merge — deletes two merged slides they do not
   * want, and appends two of their own: a summary and a thank-you. The deck is
   * 18 again, so `grew` is 6 and `added` is 6 and every clamp passes. The plan
   * is the last six slides, and two of them are the user's.
   *
   * The window is widest through the crumb, which survives the session: the
   * offer is made the next day, on a deck edited in between.
   *
   * Every merged slide carries `SSF_MERGE_RUN` in the package, written before
   * the insert where nothing can refuse it. So the run can ASK, and position
   * stops being the only evidence.
   */
  const plan = { from: 12, count: 6 };

  it("keeps only the slides that carry this run's tag", () => {
    // Four of ours, then the two the user made.
    const tags = ["r1", "r1", "r1", "r1", undefined, undefined];
    expect(provenSweep(plan, tags, "r1")).toEqual([15, 14, 13, 12]);
  });

  it("deletes highest index first, so removing one cannot shift another", () => {
    const order = provenSweep(plan, ["r1", "r1", "r1", "r1", "r1", "r1"], "r1");
    expect(order).toEqual([...order].sort((a, b) => b - a));
  });

  it("leaves a slide another run made, when this run's id is among the answers", () => {
    const tags = ["r1", "r1", "r2", "r2", "r2", "r2"];
    expect(provenSweep(plan, tags, "r2")).toEqual([17, 16, 15, 14]);
  });

  it("falls back to any slide this add-in made when the id proves nothing", () => {
    /**
     * A run recovered from a crumb carries a placeholder id — the real one was
     * never written, because the run died before it could answer. Filtering on
     * it would match nothing and take back none of the slides that are sitting
     * in the deck, which is the case recovery exists for. When the id is not
     * among the answers it cannot discriminate, so the question falls back to
     * the one that still can: did THIS ADD-IN make the slide.
     */
    const tags = ["r1", "r1", "r1", "r1", undefined, undefined];
    expect(provenSweep(plan, tags, "recovered")).toEqual([15, 14, 13, 12]);
  });

  it("takes the whole plan when the host answered nothing at all", () => {
    // No worse than before the tags were asked. A host that cannot answer is
    // not evidence that a slide is not ours — this repo's own rule is that an
    // empty read is not an empty slide.
    expect(provenSweep(plan, [undefined, undefined, undefined, undefined, undefined, undefined], "r1")).toEqual([
      17, 16, 15, 14, 13, 12,
    ]);
  });

  it("takes the whole plan when the answers do not line up with it", () => {
    // A short read is a read that failed, not a set of slides that are not
    // ours. Deleting on it would be acting on an answer nobody gave.
    expect(provenSweep(plan, ["r1", "r1"], "r1")).toEqual([17, 16, 15, 14, 13, 12]);
  });

  it("deletes nothing when every slide in the range belongs to somebody else", () => {
    // The case the whole function is for: the run's own slides are gone and
    // what is left at those positions is the user's. Nothing is offered, and
    // nothing is worse than that.
    expect(provenSweep(plan, [undefined, undefined, undefined, undefined, undefined, "x"], "r1")).toEqual([17]);
  });
});
