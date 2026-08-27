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
    dropCrumb({ deckAtStart: 12, added: 720, runId: "r12-240-abc" });
    expect(readCrumb()).toMatchObject({ deckAtStart: 12, added: 720, runId: "r12-240-abc" });
  });

  it("is gone once the slides are", () => {
    dropCrumb({ deckAtStart: 12, added: 720, runId: "r1" });
    clearCrumb();
    expect(readCrumb()).toBeUndefined();
  });

  it("refuses a crumb from a build that wrote a different shape", () => {
    // These numbers authorise deleting part of somebody's presentation. An
    // older build's record reaching the undo path as though it were sound is
    // exactly the case worth refusing.
    real.setItem(KEY, JSON.stringify({ kind: "something-else", deckAtStart: 12, added: 720 }));
    expect(readCrumb()).toBeUndefined();
  });

  it("refuses a crumb half-written by a crash", () => {
    real.setItem(KEY, '{"kind":"ssf-merge-run","deckAtStart":12,"add');
    expect(readCrumb()).toBeUndefined();
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
      expect(readCrumb(), JSON.stringify(bad)).toBeUndefined();
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
    real.setItem(KEY, JSON.stringify({ kind: "ssf-merge-run", deckAtStart: 12, added: 0, runId: "pending" }));
    expect(readCrumb()).toMatchObject({ deckAtStart: 12, added: 0 });
  });
});

describe("a store that will not cooperate", () => {
  it("does not fail a merge when there is no storage at all", () => {
    useStorage(undefined);
    expect(() => {
      dropCrumb({ deckAtStart: 12, added: 720, runId: "r1" });
    }).not.toThrow();
    expect(readCrumb()).toBeUndefined();
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
      dropCrumb({ deckAtStart: 12, added: 720, runId: "r1" });
    }).not.toThrow();
    expect(readCrumb()).toBeUndefined();
  });

  it("still READS a crumb from a store too full to write", () => {
    // The probe is a read for this reason. A full store answers no to "is there
    // room" and yes to "is there a record", and the record is the half that
    // matters — reading it needs no quota.
    real.setItem(KEY, JSON.stringify({ kind: "ssf-merge-run", deckAtStart: 12, added: 720, runId: "r1" }));
    const full = {
      getItem: (k: string) => real.getItem(k),
      setItem: () => {
        throw new Error("QuotaExceededError");
      },
      removeItem: () => undefined,
    };
    useStorage(full);
    expect(readCrumb()).toMatchObject({ added: 720 });
    expect(() => {
      dropCrumb({ deckAtStart: 1, added: 1, runId: "r2" });
    }).not.toThrow();
  });
});
