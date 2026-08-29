import { describe, expect, it } from "vitest";
import {
  PROBE_RUN_TAG,
  creationIdReading,
  deckReadVerdict,
  exportPartsVerdict,
  insertVerdict,
  insertionBlame,
  offsetVerdict,
  selectedInsertVerdict,
  substringVerdict,
  tagVerdict,
  tornInsert,
} from "../src/host/verdicts.js";
import { sweepPlan } from "../src/host/undo.js";

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

describe("whether the export drops parts the file route keeps", () => {
  const kept = {
    supported: true,
    sourceParts: 66,
    exportParts: 66,
    sourceHasAuthors: true,
    exportHasAuthors: true,
    sourceComments: 2,
    exportComments: 2,
  };

  it("says NOT ASKED when the deck has nothing to drop", () => {
    // The whole reason this arm has a control. An export with no authors part,
    // taken from a deck that never had one, is not evidence that the export
    // dropped anything — and recording it as "keeps everything" is the
    // never-asked-read-as-an-answer mistake in a new place.
    const v = exportPartsVerdict({
      ...kept,
      sourceHasAuthors: false,
      exportHasAuthors: false,
      sourceComments: 0,
      exportComments: 0,
    });
    expect(v.verdict).toBe("unknown");
    expect(v.detail).toMatch(/NOT ASKED/);
    expect(v.detail).toMatch(/deck with comments/i);
  });

  it("says NOT ASKED when the host has no such call, and says WHY", () => {
    // Below 1.10 the template is read through getFileAsync and the question
    // does not arise. A "no" here would claim a behaviour nobody exercised.
    //
    // The DETAIL is what this asserts, and that is not fussiness. Asserting
    // only `unknown` passed against a build with the branch removed — an
    // absent `sourceParts` reaches the same verdict by another route — so the
    // test proved nothing about the guard it named. "This host has no
    // exportAsBase64Presentation" and "this sheet predates the arm" are
    // different facts that send a reader to different places.
    const v = exportPartsVerdict({ supported: false });
    expect(v.verdict).toBe("unknown");
    expect(v.detail).toContain("no exportAsBase64Presentation");
  });

  it("names the authors part when it goes missing", () => {
    const v = exportPartsVerdict({ ...kept, exportHasAuthors: false, exportParts: 65 });
    expect(v.verdict).toBe("yes");
    expect(v.detail).toContain("ppt/authors.xml");
    expect(v.detail).toContain("6867");
  });

  it("counts the comment parts that went", () => {
    const v = exportPartsVerdict({ ...kept, exportComments: 0, exportParts: 64 });
    expect(v.verdict).toBe("yes");
    expect(v.detail).toContain("2 comment part(s)");
  });

  it("reports both when both go", () => {
    const v = exportPartsVerdict({ ...kept, exportHasAuthors: false, exportComments: 0, exportParts: 63 });
    expect(v.detail).toContain("ppt/authors.xml");
    expect(v.detail).toContain("comment part(s)");
  });

  it("answers NO only when the deck HAD something and the export kept it", () => {
    // The one reading that clears the API, and it is only earned by a deck
    // that could have shown the defect.
    const v = exportPartsVerdict(kept);
    expect(v.verdict).toBe("no");
    expect(v.detail).toMatch(/kept the comments/);
  });

  it("carries a throw rather than reading it as a drop", () => {
    const v = exportPartsVerdict({ supported: true, error: "GeneralException" });
    expect(v.verdict).toBe("threw");
    expect(v.detail).toContain("GeneralException");
  });

  it("says NOT ASKED for a sheet taken before the arm existed", () => {
    // Older sheets under docs/host-answers/ carry no `exportParts` at all, and
    // an absent field is not a finding.
    expect(exportPartsVerdict({}).verdict).toBe("unknown");
  });
});

describe("a partial insert read in ROWS", () => {
  /** 240 rows, three slides each — the shape the pane's summary describes. */
  const uniform = (rows: number, each = 3): number[] => Array.from({ length: rows }, () => each);

  it("says every row is whole when everything landed", () => {
    expect(tornInsert(uniform(240), 720)).toMatchObject({ complete: 240, torn: 0, absent: 0 });
  });

  it("turns 719 of 720 into one incomplete ROW", () => {
    // The whole point. "719 of 720 slides" is true and useless: every later row
    // still looks correct and the user finds it at slide 141.
    const t = tornInsert(uniform(240), 719);
    expect(t).toMatchObject({ complete: 239, torn: 1, absent: 0, firstIncomplete: 239 });
    expect(t.detail).toContain("239 of 240 row(s) landed complete");
    expect(t.detail).toContain("got 2 of its 3 slide(s)");
  });

  it("counts a row that got NOTHING apart from one that got some", () => {
    // Torn is the worse of the two, because a row with two of three slides
    // looks finished.
    const t = tornInsert(uniform(4), 6);
    expect(t).toMatchObject({ complete: 2, torn: 0, absent: 2 });
    expect(t.detail).toContain("2 row(s) got nothing");
  });

  it("reports both a torn row and the ones behind it", () => {
    const t = tornInsert(uniform(4), 7);
    expect(t).toMatchObject({ complete: 2, torn: 1, absent: 1 });
    expect(t.detail).toContain("row 3 got 1 of its 3 slide(s)");
    expect(t.detail).toContain("1 row(s) got nothing");
  });

  it("handles rows of DIFFERENT sizes, which a condition produces", () => {
    // A conditional slide leaves a row shorter, so the counts are not uniform
    // and the arithmetic cannot assume a single slides-per-row.
    const t = tornInsert([3, 1, 3, 2], 5);
    expect(t).toMatchObject({ complete: 2, torn: 1, firstIncomplete: 2 });
    expect(t.detail).toContain("row 3 got 1 of its 3 slide(s)");
  });

  it("says nothing landed when nothing landed", () => {
    expect(tornInsert(uniform(240), 0)).toMatchObject({ complete: 0, torn: 0, absent: 240 });
  });

  it("counts a one-slide-per-row merge exactly", () => {
    // No torn row is possible when a row is one slide: it either landed or it
    // did not, and reporting a tear there would be inventing one.
    const t = tornInsert(uniform(5, 1), 3);
    expect(t).toMatchObject({ complete: 3, torn: 0, absent: 2 });
  });
});

describe("undo after a torn insert", () => {
  it("takes back everything that landed, leaving no orphan", () => {
    // Checked rather than assumed: a review claimed a torn insert would leave
    // "orphan slides from a half-landed record" because sweepPlan clamps on
    // counts. It does not. `added` is the MEASURED delta, not what the plan
    // hoped for, so the sweep removes exactly what arrived — the partial row
    // included.
    const deckAtStart = 12;
    const landed = 719; // one slide short of 720
    const plan = sweepPlan({ deckAtStart, deckNow: deckAtStart + landed, added: landed });
    expect(plan).toEqual({ from: deckAtStart, count: landed });
  });
});

describe("whether a collection load of the deck answers in full", () => {
  it("will not call a small deck an answer", () => {
    // The arm this closes was live-run and reported nothing at all. When it
    // finally spoke, its first real sheet had EIGHT slides — a full read that
    // says the collection is not broken outright and says nothing whatever
    // about the ~50 ceiling office-js#4272 describes.
    const v = deckReadVerdict({
      deckSize: 8,
      items: 8,
      short: false,
      empty: false,
      canAnswerFiftyQuestion: false,
      byPosition: 8,
      prefixOk: true,
    });
    expect(v.verdict).toBe("unknown");
    expect(v.detail).toContain("NOT PUT");
    expect(v.detail).toContain("8");
  });

  it("answers yes only above the ceiling", () => {
    const v = deckReadVerdict({
      deckSize: 60,
      items: 60,
      short: false,
      empty: false,
      canAnswerFiftyQuestion: true,
      prefixOk: true,
    });
    expect(v.verdict).toBe("yes");
  });

  it("separates a bounded short read from a scrambled one", () => {
    const bounded = deckReadVerdict({ deckSize: 60, items: 50, short: true, prefixOk: true });
    const scrambled = deckReadVerdict({ deckSize: 60, items: 50, short: true, prefixOk: false });
    expect(bounded.verdict).toBe("no");
    expect(scrambled.verdict).toBe("no");
    // The difference is the whole point of the arm: a prefix-stable short read
    // refuses a block past it, a scrambled one clones slides nobody chose.
    expect(bounded.detail).toContain("DECK ORDER");
    expect(scrambled.detail).toContain("wrong slide number");
  });

  it("names the empty read as the sync-succeeded case", () => {
    const v = deckReadVerdict({ deckSize: 8, items: 0, short: true, empty: true });
    expect(v.verdict).toBe("no");
    expect(v.detail).toContain("6363");
  });

  it("says nothing when the arm was never run", () => {
    expect(deckReadVerdict({}).verdict).toBe("unknown");
    expect(deckReadVerdict({ error: "timed out" }).verdict).toBe("threw");
  });
});

describe("whether a slide insert survives a standing selection", () => {
  it("refuses to answer when nothing was selected", () => {
    // The trap this exists for: the insert lands cleanly, the sheet looks like
    // a pass, and the CONDITION was never present. The arm is read-only about
    // the selection by design, so it can only observe what the user made.
    const v = selectedInsertVerdict({ shapesSelected: 0, landed: 2, expected: 2 });
    expect(v.verdict).toBe("unknown");
    expect(v.detail).toContain("NOT ASKED");
    expect(v.detail).toContain("Re-run with a shape clicked");
  });

  it("answers no when a selection was standing and the slides landed", () => {
    const v = selectedInsertVerdict({ shapesSelected: 3, landed: 2, expected: 2 });
    expect(v.verdict).toBe("no");
    expect(v.detail).toContain("setSelectedShapes");
  });

  it("answers yes when a standing selection cost slides", () => {
    const v = selectedInsertVerdict({ shapesSelected: 3, landed: 0, expected: 2 });
    expect(v.verdict).toBe("yes");
    expect(v.detail).toContain("BLOCKS");
  });

  it("distinguishes a host that would not say from a deck with nothing selected", () => {
    // getSelectedShapes is 1.5 and the floor is 1.2, so an older host cannot
    // answer — which is a different unknown from "nothing was selected", and
    // sends the next run somewhere else.
    const v = selectedInsertVerdict({
      shapesSelected: -1,
      selectionReadError: "not supported",
      landed: 2,
      expected: 2,
    });
    expect(v.verdict).toBe("unknown");
    expect(v.detail).toContain("1.5");
    expect(v.detail).not.toContain("Re-run with a shape clicked");
  });
});

describe("an insert the deck cannot account for", () => {
  /**
   * "5 of 3 slide(s) landed, which is a partial insert" is a sentence that
   * cannot be true, and "-2 of 3" is worse. Both came out of the branch that
   * runs when the delta is neither zero nor what was expected — which named a
   * CAUSE for a condition it had not distinguished.
   *
   * `unknown` rather than `no`: slides plainly arrived. What the run cannot do
   * is say which of them are its own, and saying that is the whole answer.
   */
  it("says so when more arrived than the package held", () => {
    const v = insertVerdict({ before: 10, after: 15, expected: 3 });
    expect(v.verdict).toBe("unknown");
    expect(v.detail).toContain("grew by 5 while the package held 3");
    expect(v.detail, "named a cause it had not established").not.toContain("partial insert");
  });

  it("says so when the deck shrank across an insert", () => {
    const v = insertVerdict({ before: 10, after: 8, expected: 3 });
    expect(v.verdict).toBe("unknown");
    expect(v.detail).toContain("SHRANK by 2");
  });

  it("still grades the cases it always could", () => {
    // The other half: naming two new conditions must not blur the three that
    // were already right.
    expect(insertVerdict({ before: 10, after: 13, expected: 3 }).verdict).toBe("yes");
    expect(insertVerdict({ before: 10, after: 10, expected: 3 }).verdict).toBe("no");
    const partial = insertVerdict({ before: 10, after: 12, expected: 3 });
    expect(partial.verdict).toBe("no");
    expect(partial.detail).toContain("partial insert");
  });
});
