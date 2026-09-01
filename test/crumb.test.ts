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

  it("dates the MERGE, not the write, when a press re-writes it", () => {
    // A partial press re-drops the crumb with the count that is left, and the
    // date was re-stamped with it — so the recovery notice said "a merge from
    // <today> … and the pane closed before you could take them back" about a
    // run that was older and whose pane had not closed. The date is the
    // merge's; only the numbers change.
    // Seeded rather than written twice: two writes in the same millisecond
    // carry the same stamp, so the test would pass against the defect roughly
    // whenever the machine was quick.
    const WHEN = "2026-08-20T09:15:00.000Z";
    real.setItem(
      KEY,
      JSON.stringify({ kind: "ssf-merge-run", deckAtStart: 12, added: 6, runId: "r1", startedAt: WHEN, doc: DECK }),
    );
    dropCrumb({ deckAtStart: 12, added: 4, runId: "r1", doc: DECK, pressed: true });
    expect(readCrumb(DECK)).toMatchObject({ added: 4, pressed: true, startedAt: WHEN });

    // A DIFFERENT run is a different merge and takes its own date, whatever is
    // in the store.
    dropCrumb({ deckAtStart: 12, added: 6, runId: "r2", doc: DECK });
    expect(readCrumb(DECK)?.runId).toBe("r2");
    expect(readCrumb(DECK)?.startedAt, "a new run keeps none of the old one's").not.toBe(WHEN);
  });

  it("keeps the mark that says a press already proved these slides cannot go", () => {
    // The crumb outlives the withdrawal on purpose — it is what stops the next
    // merge overwriting a run whose slides are still in the deck — so it has to
    // carry what the press learned, or the pane offers the same dead button on
    // every open under a sentence that is not true.
    dropCrumb({ deckAtStart: 12, added: 6, runId: "r1", doc: DECK, pressed: true, unremovable: true });
    expect(readCrumb(DECK)).toMatchObject({ added: 6, pressed: true, unremovable: true });
    // An older build's crumb carries neither, and reads as neither. Written
    // under the OLD single key, which is where that build put it — and the
    // per-deck one has to be out of the way, or it answers first.
    real.clear();
    real.setItem(KEY, JSON.stringify({ kind: "ssf-merge-run", deckAtStart: 12, added: 6, runId: "r1", doc: DECK }));
    expect(readCrumb(DECK)?.unremovable).toBeUndefined();
  });

  it("refuses a stored date that is not one this build wrote", () => {
    // `Date.parse` was the first test and it is far looser than the writer:
    // "2026", "0" and "Mar 2026 junk" all parse, and the recovery notice prints
    // the first ten characters of whatever is carried — so a corrupt record
    // outlived every re-write and reached the user as a date.
    for (const junk of ["2026", "0", "1", "Mar 2026 junk", "2026-13-45T00:00:00.000Z"]) {
      real.setItem(
        KEY,
        JSON.stringify({ kind: "ssf-merge-run", deckAtStart: 12, added: 6, runId: "r1", startedAt: junk, doc: DECK }),
      );
      dropCrumb({ deckAtStart: 12, added: 4, runId: "r1", doc: DECK });
      expect(readCrumb(DECK)?.startedAt, `carried ${junk}`).not.toBe(junk);
    }
  });

  it("does not carry a date across two runs that share an id", () => {
    // Two of the pane's writes use a shared id rather than a run's own —
    // "pending", before an insert has answered, and "recovered", for a run
    // rebuilt from a crumb. Matching on the id alone made a merge inherit the
    // date of an unrelated one days earlier, which is the very thing the
    // carry-over exists to prevent.
    const WHEN = "2026-08-20T09:15:00.000Z";
    real.setItem(
      KEY,
      JSON.stringify({
        kind: "ssf-merge-run",
        deckAtStart: 18,
        added: 6,
        runId: "pending",
        startedAt: WHEN,
        doc: DECK,
      }),
    );
    dropCrumb({ deckAtStart: 40, added: 0, runId: "pending", doc: DECK });
    expect(readCrumb(DECK)?.startedAt, "a different deck size is a different run").not.toBe(WHEN);

    // And a stored date that is not one is replaced rather than carried, or a
    // corrupt record would outlive every re-write.
    real.setItem(
      KEY,
      JSON.stringify({
        kind: "ssf-merge-run",
        deckAtStart: 12,
        added: 6,
        runId: "r1",
        startedAt: "gibberish",
        doc: DECK,
      }),
    );
    dropCrumb({ deckAtStart: 12, added: 4, runId: "r1", doc: DECK });
    expect(readCrumb(DECK)?.startedAt).not.toBe("gibberish");
  });

  it("is gone once the slides are", () => {
    dropCrumb({ deckAtStart: 12, added: 720, runId: "r1", doc: DECK });
    clearCrumb(DECK);
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

  it("gives every deck a record of its own", () => {
    // ONE KEY PER DECK. A single key made every write a choice between
    // destroying another deck's record of slides still sitting in it and
    // refusing to write at all — and the refusal was its own defect: a deck
    // whose merge was never swept locked every other deck out of crash
    // recovery for the life of the browser profile.
    dropCrumb({ deckAtStart: 12, added: 6, runId: "r1", doc: DECK, unremovable: true });
    dropCrumb({ deckAtStart: 3, added: 2, runId: "r2", doc: OTHER_DECK });
    expect(readCrumb(DECK), "the first deck keeps its record").toMatchObject({ added: 6, runId: "r1" });
    expect(readCrumb(OTHER_DECK), "and the second deck gets one").toMatchObject({ added: 2, runId: "r2" });
    // Clearing one leaves the other alone.
    clearCrumb(DECK);
    expect(readCrumb(DECK)).toBeUndefined();
    expect(readCrumb(OTHER_DECK)).toMatchObject({ added: 2 });
  });

  it("does not keep a record for every deck the user has ever merged in", () => {
    // `localStorage` is not unbounded and neither is this. The oldest goes,
    // because a record whose deck has not been opened in a hundred merges is
    // the one least likely to still describe slides that are there.
    for (let i = 0; i < 12; i++) {
      real.setItem(
        `ssf-merge.run.v1:deck-${i}`,
        JSON.stringify({
          kind: "ssf-merge-run",
          deckAtStart: 1,
          added: 1,
          runId: `r${i}`,
          startedAt: `2026-08-${String(i + 1).padStart(2, "0")}T09:00:00.000Z`,
          doc: `deck-${i}`,
        }),
      );
    }
    dropCrumb({ deckAtStart: 12, added: 6, runId: "new", doc: DECK });
    const keys = Object.keys(real).filter((k) => k.startsWith("ssf-merge.run.v1:"));
    expect(keys.length, "the store grew without bound").toBeLessThanOrEqual(8);
    expect(readCrumb(DECK), "and the newest is kept").toMatchObject({ runId: "new" });
    expect(readCrumb("deck-0"), "the oldest went first").toBeUndefined();
  });

  it("does not overwrite another deck's record of slides that are still there", () => {
    // The read side refuses a stranger's crumb at length, and `clearCrumb` was
    // taught the same check — but the WRITE side had none. So opening a second
    // deck and pressing Merge erased the first deck's only record of six slides
    // still sitting in it, which is the asymmetry that makes a careful check on
    // one side worth nothing.
    dropCrumb({ deckAtStart: 12, added: 6, runId: "r1", doc: DECK });
    dropCrumb({ deckAtStart: 3, added: 0, runId: "pending-1", doc: OTHER_DECK });
    expect(readCrumb(DECK), "the first deck's slides are still recorded").toMatchObject({ added: 6 });

    // A pending marker holds nothing, so another deck may take the key — or it
    // could never be reclaimed.
    localStorage.clear();
    dropCrumb({ deckAtStart: 12, added: 0, runId: "pending-1", doc: DECK });
    dropCrumb({ deckAtStart: 3, added: 0, runId: "pending-2", doc: OTHER_DECK });
    expect(readCrumb(OTHER_DECK), "a pending marker does not lock the key").toMatchObject({ deckAtStart: 3 });
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
      clearCrumb(DECK);
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

describe("whose crumb it is, on the way OUT as well as in", () => {
  /**
   * `readCrumb` asks which deck a crumb belongs to and refuses a stranger's —
   * the docstring on `Crumb.doc` sets out why at length. `clearCrumb` asked
   * nothing at all, so a run in one deck deleted the recovery record of
   * another: the read side was careful and the write side was not.
   *
   * The store is one key for every deck the user opens, because `localStorage`
   * belongs to the add-in's ORIGIN. So this is not a corner — it is what
   * happens whenever somebody merges into a second deck while a first one has
   * a run they have not taken back.
   */
  it("will not let one deck's run delete another deck's record", () => {
    dropCrumb({ deckAtStart: 12, added: 6, runId: "r1", doc: DECK });
    // A run in a DIFFERENT deck finishes having added nothing, and clears up
    // after itself.
    clearCrumb(OTHER_DECK);
    expect(readCrumb(DECK)?.added, "deck A's six slides are still in deck A").toBe(6);
  });

  it("still clears the record of the deck it belongs to", () => {
    dropCrumb({ deckAtStart: 12, added: 6, runId: "r1", doc: DECK });
    clearCrumb(DECK);
    expect(readCrumb(DECK)).toBeUndefined();
  });

  it("clears a crumb this build cannot read rather than leaving it forever", () => {
    // A record from an older build carries no `doc`, so no caller can ever
    // match it and nothing would remove it. Refusing to clear what cannot be
    // identified would make the key unreclaimable.
    real.setItem(KEY, JSON.stringify({ kind: "ssf-merge-run", deckAtStart: 1, added: 1 }));
    clearCrumb(DECK);
    expect(real.getItem(KEY)).toBeNull();
  });

  it("does not fall over when the store refuses", () => {
    useStorage({
      getItem: () => {
        throw new Error("denied");
      },
      removeItem: () => undefined,
    });
    expect(() => clearCrumb(DECK)).not.toThrow();
  });
});
