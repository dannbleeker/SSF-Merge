import { describe, expect, it } from "vitest";
import {
  blockName,
  blockSummary,
  mergeArithmetic,
  mergeSummary,
  plural,
  slidesAdded,
  undoSummary,
} from "../src/pane/summary.js";

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
    expect(slidesAdded({ from: 4, to: 6 }, 240)).toBe(720);
  });
});

describe("what the deck will look like afterwards", () => {
  it("says where the slides land as well as how many", () => {
    // "720 slides added" answers the wrong half of the question somebody
    // hesitating over the button is actually asking.
    const s = mergeSummary({ from: 4, to: 6 }, 240, 12);
    expect(s).toContain("720 slides added after slide 12");
    expect(s).toContain("732 slides in the deck");
  });
});

describe("what an undo takes back", () => {
  it("names the slides rather than saying 'undo'", () => {
    // The pane is offering to delete part of somebody's presentation, and the
    // sentence should say so.
    expect(undoSummary(720, 732)).toBe("Remove slides 13 to 732, which this merge added.");
  });

  it("does not say 'slides 732 to 732' for a single slide", () => {
    expect(undoSummary(1, 732)).toBe("Remove slide 732, which this merge added.");
  });

  it("says there is nothing to take back rather than naming a range that is not there", () => {
    // A merge that added nothing must not offer to delete slide 13 to 12.
    expect(undoSummary(0, 12)).toBe("Nothing to take back.");
    expect(undoSummary(-3, 12)).toBe("Nothing to take back.");
  });
});
