import { describe, expect, it } from "vitest";
import {
  blockName,
  blockSummary,
  describeMerge,
  mergeArithmetic,
  mergeSummary,
  plural,
  undoIsPossible,
  undoSummary,
  type MergeReport,
} from "../src/pane/summary.js";
import { slidesPerRecord } from "../src/pane/steps.js";

describe("plurals", () => {
  it("drops the s for exactly one", () => {
    expect(plural(1, "row")).toBe("1 row");
    expect(plural(2, "row")).toBe("2 rows");
  });

  it("keeps the s for zero, which is what English does", () => {
    expect(plural(0, "slide")).toBe("0 slides");
  });
});

describe("naming the template block", () => {
  it("reads as a range when it is one", () => {
    expect(blockName({ from: 4, to: 6 })).toBe("Slides 4 to 6");
  });

  it("does not say 'slides 4 to 4'", () => {
    // The case a hand test with a three-slide template never reaches.
    expect(blockName({ from: 4, to: 4 })).toBe("Slide 4");
  });
});

describe("what the block does", () => {
  it("says the slides repeat TOGETHER, which is what makes it a block", () => {
    expect(blockSummary({ from: 4, to: 6 }, 240)).toBe("Slides 4 to 6 repeat together, 240 times.");
  });

  it("says 'once' rather than '1 times'", () => {
    expect(blockSummary({ from: 4, to: 6 }, 1)).toContain("once");
    expect(blockSummary({ from: 4, to: 6 }, 1)).not.toContain("1 times");
  });

  it("drops 'together' for a single slide, which cannot repeat with anything", () => {
    expect(blockSummary({ from: 4, to: 4 }, 240)).toBe("Slide 4 repeats 240 times.");
  });

  it("asks for data rather than guessing when there is none", () => {
    expect(blockSummary({ from: 4, to: 6 })).toContain("attach data");
  });
});

describe("the arithmetic above the merge button", () => {
  it("shows both factors so the total can be checked", () => {
    expect(mergeArithmetic({ from: 4, to: 6 }, 240)).toBe("240 rows × 3 slides");
  });

  it("multiplies both ends of the block, not the difference", () => {
    // `slidesAdded` used to do this multiplication here. It moved to
    // `plannedSlides`, which counts with the rule `buildPlan` applies, because
    // the product ignores conditions — see `test/pane-steps.test.ts`.
    expect(slidesPerRecord({ from: 4, to: 6 })).toBe(3);
  });
});

describe("what the deck will look like afterwards", () => {
  it("says where the slides land as well as how many", () => {
    // "720 slides added" answers the wrong half of the question somebody
    // hesitating over the button is actually asking.
    const s = mergeSummary(720, 12);
    // FUTURE tense: nothing has happened when this sentence is on screen.
    expect(s).toContain("720 slides will be added after slide 12");
    expect(s).toContain("732 slides in the deck");
  });

  it("says nothing about the deck's size when the host would not give one", () => {
    /**
     * `?? 0` at the call site read a refusal as an empty deck, so on a host
     * whose slide count throws — caught to `undefined` at boot and caught again
     * in `useBlock` — the sentence a user reads to decide whether to press said
     * "6 slides will be added after slide 0, leaving 6 slides in the deck" for
     * a deck with twelve slides in it. Both halves false, in the one sentence
     * that has to be true.
     */
    const s = mergeSummary(6, undefined);
    expect(s).toBe("6 slides will be added at the end of the deck.");
    // The count does not depend on the deck's size, so it is still stated. The
    // two clauses that do are dropped rather than invented.
    expect(s).not.toContain("slide 0");
    expect(s).not.toContain("in the deck.");
  });
});

describe("what an undo takes back", () => {
  it("names the slides rather than saying 'undo'", () => {
    // The pane is offering to delete part of somebody's presentation, and the
    // sentence should say so.
    expect(undoSummary(720, 732, 12)).toBe("Remove slides 13 to 732, which this merge added.");
  });

  it("does not say 'slides 732 to 732' for a single slide", () => {
    expect(undoSummary(1, 732, 731)).toBe("Remove slide 732, which this merge added.");
  });

  it("says there is nothing to take back rather than naming a range that is not there", () => {
    // A merge that added nothing must not offer to delete slide 13 to 12.
    expect(undoSummary(0, 12, 12)).toBe("Nothing to take back.");
    expect(undoSummary(-3, 12, 12)).toBe("Nothing to take back.");
  });

  it("says when a condition did nothing", () => {
    /**
     * The engine emits the slide anyway rather than hiding an authoring
     * mistake behind output that looks finished — so the merge produces the
     * RIGHT number of slides and the condition was ignored. Without this line
     * there is nothing at all to tell the user that.
     *
     * `unknownConditions` was carried on the outcome from the day the engine
     * was written and read by nothing.
     */
    expect(describeMerge({ added: 6, deckAtStart: 12, unknownConditions: ["Renewal"] })).toContain(
      "no column for Renewal, so those slides were included for every row",
    );
    // An ordinary merge stays one short sentence.
    expect(describeMerge({ added: 6, deckAtStart: 12, unknownConditions: [] })).toBe("6 slides added after slide 12.");
  });

  it("does not offer to remove slides the deck is too small to hold", () => {
    /**
     * The card computes its range backwards from the END of the deck, so a
     * deck smaller than the run's own output produced a first slide at or
     * below zero: `Remove slides -707 to 12, which this merge added.` — with a
     * button under it.
     *
     * Reachable in the product, not only in a fixture. The crash crumb offers a
     * run back when the pane reopens, and by then the user may have taken those
     * slides out by hand. `sweepPlan` already refused, so the button was safe;
     * the SENTENCE was the defect, because it said the slides were there and
     * named them.
     */
    // A deck of 12, a run that claims 720, and 12 slides there now: the run's
    // output is gone and what remains is what was there before it.
    //
    // The numbers say this differently from how they used to. "Too small to
    // hold the run's output" was measured from ZERO, because that was all the
    // card knew; it is measured from where the run started now, which is the
    // question that was always meant. A deck that began EMPTY and holds 12 of a
    // claimed 720 is a different case and gets a different answer below —
    // rightly, because nothing else has ever been in it.
    expect(undoSummary(720, 12, 12)).toBe("Nothing to take back.");
    expect(undoIsPossible(720, 12, 12)).toBe(false);
    expect(undoIsPossible(13, 12, 12)).toBe(false);
    // The boundary: a deck holding exactly the run's output and nothing else.
    expect(undoIsPossible(12, 12, 0)).toBe(true);
    expect(undoSummary(12, 12, 0)).toBe("Remove slides 1 to 12, which this merge added.");
    // And the same deck, where the run claims one more slide than ever
    // arrived. The twelve that are there can only be the run's, so they are
    // offered — `undoInsert` counts the deck again afterwards and reports what
    // actually went.
    expect(undoSummary(13, 12, 0)).toBe("Remove slides 1 to 12, which this merge added.");
  });

  it("names the range the button will actually remove, not one counted from the end", () => {
    /**
     * Three ways the deck can move between the merge and the press, and the
     * card was wrong in all three. It computed its range backwards from the END
     * of the deck; `sweepPlan` computes it from where the run started, and
     * `sweepPlan` is what the button calls.
     *
     * The button was safe throughout — it refuses what it cannot prove is the
     * run's own. What the user READ was the defect, and on a card whose whole
     * design principle is that it is offering to delete part of somebody's
     * presentation and the sentence should say which part.
     */
    // A colleague appended five slides after the merge. Slides 16 to 20 are
    // THEIRS; the merge's own are 11 to 15 and no longer identifiable by
    // position, which is why the sweep refuses.
    expect(undoSummary(5, 20, 10)).toBe("Nothing to take back.");
    expect(undoIsPossible(5, 20, 10)).toBe(false);

    // The user took three of the merged slides out by hand. Two are left, and
    // they are the last two — not a five-slide range starting three slides
    // before the merge did.
    expect(undoSummary(5, 12, 10)).toBe("Remove slides 11 to 12, which this merge added.");

    // The user took them all out. Nothing of the merge is left, and the five
    // slides at the end of the deck are the user's own.
    expect(undoSummary(5, 10, 10)).toBe("Nothing to take back.");
  });
});

describe("describeMerge says what the merge DID", () => {
  it("says a zero out loud rather than omitting it", () => {
    // The failure a first real run is likeliest to hit: a template whose
    // placeholders are spelled how its author spelled them. Every slide lands,
    // nothing is filled, and the deck delta reports a perfect success.
    const line = describeMerge({ added: 720, deckAtStart: 12, paragraphsMerged: 0 });
    expect(line).toContain("720 slides added after slide 12");
    expect(line).toMatch(/no \{\{fields\}\} were filled/i);
    expect(line).toMatch(/spelling/i);
  });

  it("counts the placeholders when the merge worked", () => {
    expect(describeMerge({ added: 6, deckAtStart: 3, paragraphsMerged: 18 })).toContain("18 placeholders filled");
  });

  it("reconciles rows against slides when some were left out", () => {
    // "8 rows" and "6 slides" are both right and a user who cannot reconcile
    // them assumes the merge lost something.
    const line = describeMerge({ added: 6, deckAtStart: 3, paragraphsMerged: 12, skippedRecords: 2 });
    // A ROW is dropped only by `onEmpty: "skip"`, when a field on one of its
    // slides has no value; `buildPlan` fills `skippedRecords` from nowhere
    // else. This said "by a condition", which sends the user to the condition
    // control, where there is nothing to find. A slide IS left out by one.
    expect(line).toContain("2 rows skipped for a blank field");
  });

  it("stays one short sentence when nothing unusual happened", () => {
    const line = describeMerge({ added: 6, deckAtStart: 3, paragraphsMerged: 12 });
    expect(line).toBe("6 slides added after slide 3 · 12 placeholders filled.");
  });

  it("says nothing about a count it was not given", () => {
    // An older outcome has no `paragraphsMerged`, and inventing "0 filled" for
    // it would report a failure that did not happen.
    expect(describeMerge({ added: 6, deckAtStart: 3 })).toBe("6 slides added after slide 3.");
  });

  it("names a chart whose own data could not be merged, and says the slides are fine", () => {
    // The chart READS correctly — the cache is what PowerPoint draws — and only
    // "Edit Data" shows the template's placeholder. Silence would leave that to
    // be found by whoever the deck is sent to.
    const line = describeMerge({ added: 6, deckAtStart: 3, paragraphsMerged: 12, workbooksUnreadable: 2 });
    expect(line).toContain("the data behind 2 charts could not be merged");
    expect(line).toContain("the slides read correctly");
  });

  it("says nothing about workbooks when every one of them merged", () => {
    const line = describeMerge({ added: 6, deckAtStart: 3, paragraphsMerged: 12, workbooksUnreadable: 0 });
    expect(line).toBe("6 slides added after slide 3 · 12 placeholders filled.");
  });

  /**
   * The other half of the merge, which said nothing at all.
   *
   * `paragraphsMerged` exists because a run that adds every slide and fills no
   * placeholder arrives looking like a success. The pictures have exactly that
   * shape and no number: the pane's pre-merge tally matches file NAMES and
   * never opens one, so a folder of renamed files passes it, places nothing,
   * and the deck grows by 720 slides either way.
   */
  const pictures = (over: Partial<NonNullable<MergeReport["pictures"]>>) => ({
    added: 6,
    deckAtStart: 3,
    paragraphsMerged: 12,
    pictures: { placed: 0, missing: [], unreadable: [], stretched: [], crowded: [], ...over },
  });

  it("counts the pictures it placed", () => {
    expect(describeMerge(pictures({ placed: 4 }))).toContain("4 pictures placed");
  });

  it("says a zero out loud, the way the placeholder count does", () => {
    const line = describeMerge(pictures({ missing: ["Photo"] }));
    expect(line).toContain("no pictures were placed");
  });

  it("names a file that is not a picture, which nothing before the merge can catch", () => {
    const line = describeMerge(pictures({ unreadable: ["Photo"] }));
    expect(line).toContain("the file for Photo is not a picture this add-in can read");
  });

  it("names a field whose shape was already taken", () => {
    // A shape has one fill. The second field in it was dropped in silence.
    expect(describeMerge(pictures({ placed: 1, crowded: ["Logo"] }))).toContain("Logo had nowhere to go");
  });

  it("names a picture it had to squash", () => {
    // The shape inherits its size from a layout placeholder, so there was no
    // ratio to fit to. It reads as a broken image and is not one.
    expect(describeMerge(pictures({ placed: 1, stretched: ["Photo"] }))).toContain("Photo was stretched");
  });

  it("does not name a missing picture, which is documented behaviour", () => {
    // `missing` counts an EMPTY cell as well as a file nobody supplied, and an
    // empty cell keeps its placeholder by design — exactly as a text field
    // with no column does. The count carries that case without accusing
    // anybody.
    const line = describeMerge(pictures({ placed: 2, missing: ["Photo"] }));
    expect(line).toContain("2 pictures placed");
    expect(line).not.toContain("missing");
  });

  it("stays silent on a merge that had no pictures in it", () => {
    const line = describeMerge(pictures({}));
    expect(line).toBe("6 slides added after slide 3 · 12 placeholders filled.");
  });

  it("says nothing about pictures for an outcome that carries none", () => {
    expect(describeMerge({ added: 6, deckAtStart: 3, paragraphsMerged: 12 })).toBe(
      "6 slides added after slide 3 · 12 placeholders filled.",
    );
  });

  /**
   * The third half-reported outcome, and the one whose failure is invisible.
   *
   * `mergeChartNumbers` writes nothing when a value cell's placeholder does
   * not resolve to a number — deliberately, because guessing zero draws a
   * chart the data never said. So the point keeps the TEMPLATE's number under
   * the merged label: a chart that is wrong and looks right. It counted those
   * refusals for a reader that did not exist.
   */
  const values = (filled: number, refused: number, unreadable = 0, unplotted = 0) => ({
    added: 6,
    deckAtStart: 3,
    paragraphsMerged: 12,
    chartValues: { filled, refused, unreadable, unplotted },
  });

  it("says when a chart's data could not be read at all", () => {
    /**
     * The third outcome, which had no clause because it had no count. A series
     * whose range cannot be read, or which names a sheet the workbook does not
     * declare, was abandoned with a bare `continue` — no fill and no refusal —
     * so the pane was handed `{filled: 0, refused: 0}` and said nothing.
     *
     * The comment above this block reasoned from exactly that pair: "every
     * value placeholder either fills or refuses, so a zero here is ALWAYS a
     * refusal". It was not, and the case it missed is the one where the user
     * gets no signal of any kind.
     */
    expect(describeMerge(values(0, 0, 1))).toContain("could not be read");
    // And it says so beside the ones that did work, rather than instead of them.
    const both = describeMerge(values(4, 0, 2));
    expect(both).toContain("4 chart values filled");
    expect(both).toContain("could not be read");
  });

  it("stays quiet when every series was read", () => {
    expect(describeMerge(values(9, 0))).not.toContain("could not be read");
  });

  it("says when a value reached the data sheet but not the chart", () => {
    /**
     * The third silent outcome, and the only one with a remedy the user can
     * act on: a chart's cached point list is sparse, so a writer omits the
     * point for a cell it has no number for — which is the cell a placeholder
     * was typed into. The merge fills the data sheet and has nowhere to put the
     * value in the chart.
     *
     * So the sentence says where the value IS, not only where it is not.
     */
    const line = describeMerge(values(0, 0, 0, 2));
    expect(line).toContain("data sheet");
    expect(line, "and what to do about it").toContain("Edit Data");
  });

  it("counts the chart values it filled", () => {
    expect(describeMerge(values(9, 0))).toContain("9 chart values filled");
  });

  it("names the ones that did not read as numbers, and what the chart shows instead", () => {
    const line = describeMerge(values(6, 2));
    expect(line).toContain("2 chart values did not read as a number");
    expect(line).toContain("those points still show the template's");
  });

  it("puts one refusal in the singular, because the sentence names what it shows", () => {
    expect(describeMerge(values(0, 1))).toContain("that point still shows the template's");
  });

  it("does NOT say a zero here, because the refusal already says it", () => {
    // The asymmetry with the pictures is deliberate. Every value placeholder
    // either fills or refuses, so a zero here is always a refusal and the
    // clause below names the count — "no chart values were filled · 3 chart
    // values did not read as a number" is one fact reported twice. A picture
    // can be absent for a reason the sentence deliberately does not name, so
    // there the zero is the only thing said.
    const line = describeMerge(values(0, 3));
    expect(line).not.toContain("no chart values were filled");
    expect(line).toContain("3 chart values did not read as a number");
  });

  it("stays silent on a merge with no chart values in it", () => {
    expect(describeMerge(values(0, 0))).toBe("6 slides added after slide 3 · 12 placeholders filled.");
  });
});
