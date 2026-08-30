// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { clearCrumb, dropCrumb, readCrumb } from "../src/pane/crumb.js";

const KEY = "ssf-merge.run.v1";
const real = globalThis.localStorage;

/** Put a hostile Storage in place, and take it away again. */
function useStorage(fake: Partial<Storage> | undefined): void {
  Object.defineProperty(globalThis, "localStorage", {
    value: fake,
    configurable: true,
    writable: true,
  });
}

const DECK = "https://example-my.sharepoint.com/personal/x/Documents/deck.pptx";
const OTHER_DECK = "https://example-my.sharepoint.com/personal/x/Documents/something-else.pptx";

beforeEach(() => {
  useStorage(real);
  real.clear();
});

afterEach(() => useStorage(real));

describe("the numbers an undo cannot be done without", () => {
  it("survives the pane being thrown away", () => {
    // The whole point. `deckAtStart` lives in a module variable, so a tab that
    // dies during a merge leaves the deck holding the slides and nothing able
    // to say how many or where they start.
    dropCrumb({ deckAtStart: 12, added: 720, runId: "r12-240-abc", doc: DECK });
    expect(readCrumb(DECK)).toMatchObject({ deckAtStart: 12, added: 720, runId: "r12-240-abc" });
  });

  it("is gone once the slides are", () => {
    dropCrumb({ deckAtStart: 12, added: 720, runId: "r1", doc: DECK });
    clearCrumb();
    expect(readCrumb(DECK)).toBeUndefined();
  });

  it("refuses a crumb belonging to a DIFFERENT document", () => {
    // The defect the real-host round of 2026-08-30 found. `localStorage` is
    // scoped to the add-in's origin, not to the deck, so one deck's interrupted
    // run was read while another was open — a fresh, never-merged deck was told
    // "a merge from 2026-08-30 added 6 slide(s)".
    //
    // That instance was only alarming, because the other deck had not grown and
    // `sweepPlan` refused. It is not safe in general: every clamp there compares
    // SIZES, so a stranger deck holding exactly deckAtStart + added slides
    // passes all of them and yields a plan to delete six slides this run never
    // created. Identity is the question none of those clamps can ask.
    dropCrumb({ deckAtStart: 3, added: 6, runId: "r3-3-huqtal", doc: DECK });
    expect(readCrumb(DECK), "the deck it was written on still answers").toMatchObject({ added: 6 });
    expect(readCrumb(OTHER_DECK)).toBeUndefined();
  });

  it("refuses when the host will not name the document", () => {
    // "" is what `documentKey()` answers for a host that will not say, and it
    // cannot be told apart from a match by guessing — so both sides must be
    // known. The cost is that dead-tab recovery is not offered there, which is
    // the direction `sweepPlan` already takes when it cannot prove a slide is
    // the run's own.
    dropCrumb({ deckAtStart: 3, added: 6, runId: "r1", doc: DECK });
    expect(readCrumb("")).toBeUndefined();
  });

  it("refuses a crumb from a build that named no document at all", () => {
    // An older build's record: the shape validates, the numbers are sound, and
    // it still cannot say which deck it belongs to.
    real.setItem(KEY, JSON.stringify({ kind: "ssf-merge-run", deckAtStart: 3, added: 6, runId: "r1" }));
    expect(readCrumb(DECK)).toBeUndefined();
  });

  it("refuses a crumb from a build that wrote a different shape", () => {
    // These numbers authorise deleting part of somebody's presentation. An
    // older build's record reaching the undo path as though it were sound is
    // exactly the case worth refusing.
    real.setItem(KEY, JSON.stringify({ kind: "something-else", deckAtStart: 12, added: 720 }));
    expect(readCrumb(DECK)).toBeUndefined();
  });

  it("refuses a crumb half-written by a crash", () => {
    real.setItem(KEY, '{"kind":"ssf-merge-run","deckAtStart":12,"add');
    expect(readCrumb(DECK)).toBeUndefined();
  });

  it("refuses numbers a sweep could not be clamped by", () => {
    // Both, or neither. "I could not read it" and "the deck was this big" must
    // never be confused on a path that deletes slides.
    for (const bad of [
      { kind: "ssf-merge-run", deckAtStart: 12 },
      { kind: "ssf-merge-run", added: 720 },
      { kind: "ssf-merge-run", deckAtStart: -1, added: 720 },
      { kind: "ssf-merge-run", deckAtStart: 12, added: -1 },
      { kind: "ssf-merge-run", deckAtStart: 1.5, added: 720 },
    ]) {
      real.setItem(KEY, JSON.stringify(bad));
      expect(readCrumb(DECK), JSON.stringify(bad)).toBeUndefined();
    }
  });

  it("KEEPS a crumb written before the insert, which is the one it exists for", () => {
    /**
     * `merge()` writes `{added: 0}` before handing the package to PowerPoint,
     * because a tab that dies during that call never comes back to write the
     * real number — and that is the whole reason this file exists. Refusing
     * zero on the way out made the crumb write-only in exactly that window:
     * readable only after the run it was insurance against had succeeded.
     *
     * Zero authorises nothing. `sweepPlan` refuses a count of zero, so what the
     * caller gets from this is a sentence and the deck's size before the run,
     * never a delete.
     */
    real.setItem(
      KEY,
      JSON.stringify({ kind: "ssf-merge-run", deckAtStart: 12, added: 0, runId: "pending", doc: DECK }),
    );
    expect(readCrumb(DECK)).toMatchObject({ deckAtStart: 12, added: 0 });
  });
});

describe("a store that will not cooperate", () => {
  it("does not fail a merge when there is no storage at all", () => {
    useStorage(undefined);
    expect(() => {
      dropCrumb({ deckAtStart: 12, added: 720, runId: "r1", doc: DECK });
    }).not.toThrow();
    expect(readCrumb(DECK)).toBeUndefined();
    expect(() => {
      clearCrumb();
    }).not.toThrow();
  });

  it("does not fail a merge when policy blocks the store", () => {
    // A task pane is a third-party frame, and a store that exists and throws on
    // ACCESS is a real configuration rather than a corner case.
    const hostile = {
      getItem: () => {
        throw new Error("blocked by policy");
      },
      setItem: () => {
        throw new Error("blocked by policy");
      },
      removeItem: () => {
        throw new Error("blocked by policy");
      },
    };
    useStorage(hostile);
    expect(() => {
      dropCrumb({ deckAtStart: 12, added: 720, runId: "r1", doc: DECK });
    }).not.toThrow();
    expect(readCrumb(DECK)).toBeUndefined();
  });

  it("still READS a crumb from a store too full to write", () => {
    // The probe is a read for this reason. A full store answers no to "is there
    // room" and yes to "is there a record", and the record is the half that
    // matters — reading it needs no quota.
    real.setItem(KEY, JSON.stringify({ kind: "ssf-merge-run", deckAtStart: 12, added: 720, runId: "r1", doc: DECK }));
    const full = {
      getItem: (k: string) => real.getItem(k),
      setItem: () => {
        throw new Error("QuotaExceededError");
      },
      removeItem: () => undefined,
    };
    useStorage(full);
    expect(readCrumb(DECK)).toMatchObject({ added: 720 });
    expect(() => {
      dropCrumb({ deckAtStart: 1, added: 1, runId: "r2", doc: DECK });
    }).not.toThrow();
  });
});
