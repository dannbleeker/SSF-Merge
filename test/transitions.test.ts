/**
 * The pane's multi-field state rules.
 *
 * A change that sets one field is readable where it happens. One that clears
 * six is a RULE, and a rule written out at each call site drifts — this one had,
 * four ways.
 */
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { blockMoved, blockTyped, dataChanged } from "../src/pane/transitions.js";
import type { PaneState } from "../src/pane/steps.js";

const FULL: PaneState = {
  block: { from: 3, to: 5 },
  fields: ["Name", "Photo"],
  imageFields: ["Photo"],
  conditions: { 5: "Renewal" },
  slideFields: [["Name"], ["Photo"]],
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

  /**
   * Everything a template read produces, cleared together.
   *
   * Named from `BlockReport` rather than listed here, because listing is how
   * the last one was missed: `slideFields` was added to the read, to the
   * outcome and to the state, and not to this rule — so moving the block left
   * the pane counting rows against the OLD block's per-slide fields, dropping
   * a row for a blank in a `{{Note}}` on a slide the state no longer named.
   *
   * A fourth field on that report joins this automatically.
   */
  it("drops everything the template read produced, whatever it is called", () => {
    const source = readFileSync("src/office/merge.ts", "utf8");
    const at = source.indexOf("export interface BlockReport");
    const body = source.slice(at, source.indexOf("\n}", at));
    const produced = [...body.matchAll(/^ {2}(\w+)\??:/gm)]
      .map((m) => m[1] ?? "")
      // Not read off the slides: one is the verdict and one is its sentence.
      .filter((name) => name !== "ok" && name !== "detail");

    expect(produced, "found the report's fields at all").toContain("slideFields");
    for (const name of produced) {
      const before = (FULL as unknown as Record<string, unknown>)[name];
      expect(before, `${name} is set on the state going in, or this proves nothing`).toBeTruthy();
      const value = (after as unknown as Record<string, unknown>)[name];
      const gone = value === undefined || (Array.isArray(value) && value.length === 0);
      expect(gone, `blockMoved left ${name} standing, read off slides it no longer names`).toBe(true);
    }
  });

  it("drops the conditions, which are keyed by slide NUMBER", () => {
    // "Slide 5 only when Renewal" is about the fifth slide of the deck. A block
    // starting one slide later applies it to a different slide.
    expect(after.conditions).toBeUndefined();
  });

  it("disarms a finished run and its note, and KEEPS what it added", () => {
    expect(after.changedSinceMerge, "the button is armed for a merge this state no longer describes").toBe(true);
    expect(after.fieldNote, "reports an insert into slides this state no longer names").toBeUndefined();
    // Moving the block does not take the last run's slides out of the deck,
    // and `added` is what the undo card is drawn from. Clearing it withdrew
    // the only offer to remove them while they were still there.
    expect(after.added, "still in the deck, so still offered back").toBe(12);
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

describe("when the slide numbers are typed in", () => {
  /**
   * `blockMoved` is right for its own contract; this is about WHEN it fires. It
   * ran on every keystroke, so a user who set "slide 5 only when Renewal", went
   * back to check the slide numbers, retyped the same last slide and walked
   * forward again found the conditions gone — silently, with the merge button
   * quietly offering more slides than they had asked for.
   */
  it("keeps the conditions when the numbers still name the same block", () => {
    const same = blockTyped(FULL, { from: "3", to: "5" });
    expect(same.conditions, "the block did not move, so nothing became stale").toEqual({ 5: "Renewal" });
    // Everything read off the SLIDES still goes: this pane has no way to know
    // they were not edited between the read and the keystroke.
    expect(same.fields).toEqual([]);
    expect(same.slideFields).toBeUndefined();
  });

  it("keeps them while a box is empty, which is incomplete rather than different", () => {
    // The route the user actually takes: clear the box, then retype. Dropping
    // them here would lose them before the second keystroke could keep them.
    const clearing = blockTyped(FULL, { from: "3", to: "" });
    expect(clearing.conditions).toEqual({ 5: "Renewal" });
    expect(blockTyped(clearing, { from: "3", to: "5" }).conditions).toEqual({ 5: "Renewal" });
  });

  it("drops them once the block has genuinely moved, across the empty box too", () => {
    // The case the anchor exists to get right: clear a box, then type a
    // DIFFERENT number. Without it the second keystroke reads as "no block
    // before, no move", and conditions keyed to slide 5 survive onto a block
    // that no longer contains it.
    const clearing = blockTyped(FULL, { from: "3", to: "" });
    expect(blockTyped(clearing, { from: "3", to: "9" }).conditions).toBeUndefined();
  });

  it("drops the anchor with the conditions, so a later keystroke cannot revive them", () => {
    expect(blockMoved(FULL).conditionsFor).toBeUndefined();
    const moved = blockTyped(FULL, { from: "3", to: "9" });
    expect(moved.conditionsFor).toBeUndefined();
  });

  it("drops them the moment the numbers name different slides", () => {
    // Which is the rule `blockMoved` exists for: a condition is keyed by slide
    // NUMBER, so a block starting one slide later applies it to another slide.
    expect(blockTyped(FULL, { from: "3", to: "9" }).conditions).toBeUndefined();
    expect(blockTyped(FULL, { from: "4", to: "5" }).conditions).toBeUndefined();
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

  it("disarms a finished run and its note, and KEEPS what it added", () => {
    expect(after.changedSinceMerge).toBe(true);
    expect(after.fieldNote).toBeUndefined();
    expect(after.added, "a new paste does not remove the slides the last run added").toBe(12);
  });
});
