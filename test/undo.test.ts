import { describe, expect, it } from "vitest";
import { sweepPlan } from "../src/host/undo.js";

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
