/**
 * The pane's multi-field state rules.
 *
 * A change that sets one field is readable where it happens. One that clears
 * six is a RULE, and a rule written out at each call site drifts — this one had,
 * four ways.
 */
import { describe, expect, it } from "vitest";
import { blockMoved, dataChanged } from "../src/pane/transitions.js";
import type { PaneState } from "../src/pane/steps.js";

const FULL: PaneState = {
  block: { from: 3, to: 5 },
  fields: ["Name", "Photo"],
  imageFields: ["Photo"],
  conditions: { 5: "Renewal" },
  added: 12,
  fieldNote: "{{Name}} put on the slide.",
  columns: ["Name"],
  rows: 4,
  previewing: false,
  deckSize: 40,
  excluded: [2],
  rowSearch: "ada",
  notice: "something the host said",
  draft: { from: "3", to: "5" },
};

describe("when the block moves", () => {
  const after = blockMoved(FULL);

  it("drops everything read off the old slides", () => {
    expect(after.block).toBeUndefined();
    expect(after.fields).toEqual([]);
    // Read off the block's slides in the same pass as `fields`, and added
    // after the other three call sites were written — so no path cleared it.
    expect(after.imageFields).toEqual([]);
  });

  it("drops the conditions, which are keyed by slide NUMBER", () => {
    // "Slide 5 only when Renewal" is about the fifth slide of the deck. A block
    // starting one slide later applies it to a different slide.
    expect(after.conditions).toBeUndefined();
  });

  it("disarms a finished run and its note", () => {
    expect(after.added, "the button stays armed for a merge this state no longer describes").toBeUndefined();
    expect(after.fieldNote, "reports an insert into slides this state no longer names").toBeUndefined();
  });

  it("leaves alone what the block does not decide", () => {
    // The half that keeps this from becoming a reset button.
    expect(after.deckSize).toBe(40);
    expect(after.rows).toBe(4);
    expect(after.excluded).toEqual([2]);
    expect(after.draft).toEqual({ from: "3", to: "5" });
    // Two callers set a notice of their own, so clearing it here would make the
    // order of two statements decide whether the user sees the sentence.
    expect(after.notice).toBe("something the host said");
  });
});

describe("when new data arrives", () => {
  const after = dataChanged(FULL);

  it("drops the row filter, which was about the old rows", () => {
    // Row 7 of the old paste is not row 7 of the new one.
    expect(after.excluded).toBeUndefined();
    expect(after.rowSearch).toBeUndefined();
  });

  it("keeps the conditions, which are about the template", () => {
    // A column the new data lacks is REPORTED, not quietly dropped.
    expect(after.conditions).toEqual({ 5: "Renewal" });
    expect(after.block).toEqual({ from: 3, to: 5 });
    expect(after.fields).toEqual(["Name", "Photo"]);
  });

  it("disarms a finished run and its note", () => {
    expect(after.added).toBeUndefined();
    expect(after.fieldNote).toBeUndefined();
  });
});
