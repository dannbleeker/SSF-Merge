import { describe, expect, it } from "vitest";
import {
  PROBE_RUN_TAG,
  creationIdReading,
  insertVerdict,
  insertionBlame,
  offsetVerdict,
  substringVerdict,
  sweepPlan,
  tagVerdict,
} from "../src/host/verdicts.js";

describe("insertVerdict", () => {
  it("reads a matching delta as success", () => {
    expect(insertVerdict({ before: 3, after: 5, expected: 2 }).verdict).toBe("yes");
  });

  it("calls a silent drop a failure, not a success", () => {
    // This host family has accepted slide adds, deck inserts and tag writes and
    // performed none of them. The delta is the evidence, never the absence of
    // an error.
    const v = insertVerdict({ before: 3, after: 3, expected: 2 });
    expect(v.verdict).toBe("no");
    expect(v.detail).toContain("silently");
  });

  it("separates a partial insert from a refusal", () => {
    const v = insertVerdict({ before: 3, after: 4, expected: 2 });
    expect(v.verdict).toBe("no");
    expect(v.detail).toContain("partial");
  });

  it("reports a throw as its own outcome rather than as a no", () => {
    expect(insertVerdict({ before: 3, after: 3, expected: 2, error: "InvalidArgument" }).verdict).toBe("threw");
  });
});

describe("creationIdReading", () => {
  const ok = insertVerdict({ before: 0, after: 2, expected: 2 });
  const dropped = insertVerdict({ before: 0, after: 0, expected: 2 });
  const threw = insertVerdict({ before: 0, after: 0, expected: 2, error: "InvalidArgument" });

  it("confirms the mechanism only when the two arms disagree", () => {
    expect(creationIdReading(ok, threw)).toContain("CONFIRMED");
  });

  it("says the bug does not reproduce when both arms land", () => {
    const reading = creationIdReading(ok, ok);
    expect(reading).toContain("does not reproduce");
    expect(reading).toContain("Keep the rewrite");
  });

  it("refuses to read anything into creation ids when neither arm lands", () => {
    // Asking only the fresh-id arm cannot tell "the bug is absent" from "this
    // host refuses every insert". That is why there are two arms.
    const reading = creationIdReading(dropped, dropped);
    expect(reading).toContain("BLOCKING");
    expect(reading).not.toContain("CONFIRMED");
  });

  it("refuses to conclude anything from the inverted result", () => {
    expect(creationIdReading(dropped, ok)).toContain("re-run");
  });
});

describe("substringVerdict", () => {
  const base = { before: "Hello NAME here", want: "Hello Ada here" };

  it("passes only when the text is right and the styling survived", () => {
    expect(substringVerdict({ ...base, after: "Hello Ada here", boldAfter: true }).verdict).toBe("yes");
  });

  it("fails when the text is right but the run was flattened", () => {
    const v = substringVerdict({ ...base, after: "Hello Ada here", boldAfter: false });
    expect(v.verdict).toBe("no");
    expect(v.detail).toContain("flattens");
  });

  it("blames the offsets, not the formatting, when the text is wrong", () => {
    const v = substringVerdict({ ...base, after: "HelloAda  here", boldAfter: true });
    expect(v.verdict).toBe("no");
    expect(v.detail).toContain("offsets");
  });

  it("says unknown rather than yes when the host would not report the formatting", () => {
    expect(substringVerdict({ ...base, after: "Hello Ada here" }).verdict).toBe("unknown");
  });
});

describe("offsetVerdict", () => {
  it("recognises independent offsets", () => {
    expect(offsetVerdict("AAA-BBB", "AAA-BBB", "AAA-XXX").verdict).toBe("yes");
  });

  it("recognises shifted offsets", () => {
    expect(offsetVerdict("AAA-XXX", "AAA-BBB", "AAA-XXX").verdict).toBe("no");
  });

  it("refuses to pick a model that predicts neither result", () => {
    const v = offsetVerdict("something else", "AAA-BBB", "AAA-XXX");
    expect(v.verdict).toBe("unknown");
    expect(v.detail).toContain("right to left");
  });
});

describe("sweepPlan", () => {
  it("removes only what this run added", () => {
    expect(sweepPlan({ deckAtStart: 10, deckNow: 14, added: 4 })).toEqual({ from: 10, count: 4 });
  });

  it("never removes more than the deck actually grew", () => {
    // The probe thinks it added six; the deck only grew by two. Believing the
    // probe would reach four slides back into the user's own content.
    expect(sweepPlan({ deckAtStart: 10, deckNow: 12, added: 6 })).toEqual({ from: 10, count: 2 });
  });

  it("never removes more than the run claims, even if the deck grew more", () => {
    // Somebody else added a slide while the probe ran. It is not ours.
    expect(sweepPlan({ deckAtStart: 10, deckNow: 20, added: 3 })).toEqual({ from: 17, count: 3 });
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

describe("reading back a tag written into the package", () => {
  const landed = { insertLanded: 2 };

  it("says NOT ASKED when the slide that would carry the tag never landed", () => {
    // The defect this function was written for. The probe reads the tag off the
    // LAST slide in the deck; when the insert threw, that is a slide the user
    // owns and has never carried our tag. The first real sheet reported
    // "the metadata scheme needs rethinking" on exactly that read.
    const v = tagVerdict({ insertLanded: 0 });
    expect(v.verdict).toBe("unknown");
    expect(v.detail).toContain("NOT ASKED");
  });

  it("still says NOT ASKED when the read ALSO threw", () => {
    // Order matters: a throw on a question that was never put is not a fact
    // about the host either.
    expect(tagVerdict({ insertLanded: 0, error: "InvalidArgument" }).verdict).toBe("unknown");
  });

  it("reports a missing tag as NO once the slide really did land", () => {
    expect(tagVerdict({ ...landed }).verdict).toBe("no");
  });

  it("reports the tag the probe writes as yes", () => {
    expect(tagVerdict({ ...landed, value: PROBE_RUN_TAG }).verdict).toBe("yes");
  });

  it("refuses to call a value nothing wrote an answer", () => {
    expect(tagVerdict({ ...landed, value: "something else" }).verdict).toBe("unknown");
  });

  it("reports a throw as a throw", () => {
    expect(tagVerdict({ ...landed, error: "GeneralException" }).verdict).toBe("threw");
  });
});

describe("whose fault a refused insert is", () => {
  it("blames US when the host took its own deck and refused ours", () => {
    expect(insertionBlame("threw", "yes")).toContain("OURS");
  });

  it("blames THE HOST when it refused the deck it wrote itself", () => {
    expect(insertionBlame("threw", "threw")).toContain("THE HOST");
  });

  it("refuses to blame anyone when the control never ran", () => {
    // The state the first real sheet was in. Without the control, InvalidArgument
    // is equally our package and this host, and those are opposite conclusions.
    expect(insertionBlame("threw", "unknown")).toContain("CANNOT TELL");
  });

  it("does not ask the question at all once our own insert worked", () => {
    for (const self of ["yes", "no", "threw", "unknown"] as const) {
      expect(insertionBlame("yes", self)).toContain("works");
    }
  });
});

describe("an insert that raised and landed anyway", () => {
  it("reads the DELTA, not the error, when everything asked for arrived", () => {
    // The third real sheet: a 30-second budget expired on an insert whose deck
    // delta was exactly the two slides requested. Reading the error as decisive
    // produced three false statements downstream — that our package was
    // refused, that the collision arm disagreed with the fresh one, and that
    // the theme was the difference.
    const v = insertVerdict({ before: 2, after: 4, expected: 2, error: "gave up waiting for: inserting a deck" });
    expect(v.verdict).toBe("yes");
    expect(v.landed).toBe(2);
    expect(v.detail).toContain("stopped waiting");
  });

  it("keeps a late landing out of the blame arm", () => {
    // The cascade is the point: one misread arm made insertionBlame accuse the
    // package writer of a refusal that never happened.
    const fresh = insertVerdict({ before: 2, after: 4, expected: 2, error: "gave up waiting" });
    expect(insertionBlame(fresh.verdict, "yes")).toContain("works");
  });

  it("lets a late landing carry the tag question, which depends on it", () => {
    const fresh = insertVerdict({ before: 2, after: 4, expected: 2, error: "gave up waiting" });
    expect(tagVerdict({ value: PROBE_RUN_TAG, insertLanded: fresh.landed }).verdict).toBe("yes");
  });

  it("still calls a raise that landed NOTHING a throw", () => {
    const v = insertVerdict({ before: 2, after: 2, expected: 2, error: "InvalidArgument" });
    expect(v.verdict).toBe("threw");
    expect(v.landed).toBe(0);
  });

  it("still calls a raise that landed SOME of it a throw, and says how many", () => {
    // A partial landing after a raise is not a success, and hiding the count
    // would make it look like nothing happened when a slide is really there.
    const v = insertVerdict({ before: 2, after: 3, expected: 2, error: "InvalidArgument" });
    expect(v.verdict).toBe("threw");
    expect(v.detail).toContain("1 slide(s) landed anyway");
  });
});
