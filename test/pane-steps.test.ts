import { describe, expect, it } from "vitest";
import { EMPTY, STEPS, blockedReason, primary, slidesPerRecord, statusOf, unmatchedFields } from "../src/pane/steps.js";
import type { PaneState } from "../src/pane/steps.js";

const ready: PaneState = {
  block: { from: 4, to: 6 },
  fields: ["First", "Last"],
  columns: ["First", "Last", "Email"],
  rows: 240,
  previewing: false,
};

describe("how many slides a record produces", () => {
  it("counts both ends of the block", () => {
    // Slides 4 to 6 is THREE slides. The classic off-by-one, and it multiplies
    // by the row count into the number on the button.
    expect(slidesPerRecord({ from: 4, to: 6 })).toBe(3);
  });

  it("is one for a single-slide template", () => {
    expect(slidesPerRecord({ from: 4, to: 4 })).toBe(1);
  });
});

describe("what blocks a step", () => {
  it("lets the user start on the template with nothing chosen", () => {
    expect(blockedReason(EMPTY, "template")).toBeNull();
  });

  it("blocks every later step until a block is chosen", () => {
    for (const step of ["fields", "preview", "merge"] as const) {
      expect(blockedReason(EMPTY, step), step).toContain("repeat");
    }
  });

  it("blocks preview and merge until data is attached", () => {
    const noData: PaneState = { block: { from: 4, to: 6 }, fields: [], previewing: false };
    expect(blockedReason(noData, "fields")).toBeNull();
    expect(blockedReason(noData, "preview")).toContain("data");
    expect(blockedReason(noData, "merge")).toContain("data");
  });

  it("NAMES the placeholders that have no column", () => {
    // A count alone sends the user back through every slide looking for it.
    const missing: PaneState = { ...ready, fields: ["First", "Nickname", "Badge"] };
    const why = blockedReason(missing, "merge") ?? "";
    expect(why).toContain("Nickname");
    expect(why).toContain("Badge");
    expect(why).not.toContain("First");
  });

  it("refuses to merge while a preview is still on the slide", () => {
    // The preview REPLACES the template's text and puts it back afterwards.
    // Merging mid-preview would take one row's values as the template.
    expect(blockedReason({ ...ready, previewing: true }, "merge")).toContain("preview");
  });

  it("lets the merge through once everything is answered", () => {
    expect(blockedReason(ready, "merge")).toBeNull();
  });

  it("treats a field with no data attached as not yet a problem", () => {
    // Before data arrives every placeholder is unmatched, and saying so would
    // be shouting at the user about a step they have not reached.
    const noData: PaneState = { block: { from: 4, to: 6 }, fields: ["First"], previewing: false };
    expect(unmatchedFields(noData)).toEqual([]);
  });
});

describe("the one primary button", () => {
  it("names what it does WITH THE NUMBER IN IT", () => {
    // 240 rows x 3 slides. The label is the statement a user can check against
    // the deck in front of them, where "Merge" is only a promise.
    expect(primary(ready, "merge").label).toBe("Add 720 slides");
  });

  it("says slide, singular, when it would add exactly one", () => {
    expect(primary({ ...ready, rows: 1, block: { from: 4, to: 4 } }, "merge").label).toBe("Add 1 slide");
  });

  it("cannot be pressed while the merge is blocked", () => {
    expect(primary({ ...ready, previewing: true }, "merge").enabled).toBe(false);
    expect(primary(EMPTY, "merge").enabled).toBe(false);
  });

  it("offers to put the template back while previewing, not to preview again", () => {
    expect(primary({ ...ready, previewing: true }, "preview").label).toContain("back");
  });

  it("gives every step exactly one label, and never an empty one", () => {
    for (const step of STEPS) {
      for (const state of [EMPTY, ready, { ...ready, previewing: true }]) {
        expect(primary(state, step).label.length, `${step}`).toBeGreaterThan(0);
      }
    }
  });
});

describe("what the step rail shows", () => {
  it("marks the step the user is on as current", () => {
    expect(statusOf(ready, "fields", "fields")).toBe("current");
  });

  it("marks an earlier answered step as done", () => {
    expect(statusOf(ready, "template", "fields")).toBe("done");
  });

  it("never marks a LATER step as done, however answered it is", () => {
    // Otherwise a fully-filled state paints all four ticks before the user has
    // been anywhere, and the rail stops meaning progress.
    expect(statusOf(ready, "merge", "fields")).toBe("waiting");
  });
});
