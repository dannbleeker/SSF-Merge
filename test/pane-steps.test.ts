import { describe, expect, it } from "vitest";
import {
  EMPTY,
  STEPS,
  blockedReason,
  chosenBlock,
  firstIncludedRow,
  firstRowOnly,
  includedCount,
  includedRecords,
  nextStep,
  primary,
  readBlockDraft,
  readPastedTable,
  rowLabel,
  slidesPerRecord,
  statusOf,
  unmatchedFields,
  visibleRows,
} from "../src/pane/steps.js";
import type { PaneState, StepId } from "../src/pane/steps.js";

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

  it("refuses to merge while a preview is still in the deck", () => {
    // A preview is one row already inserted. Merging on top of it would leave
    // the user with 241 sets of slides where they asked for 240, and the
    // preview's own slides are indistinguishable from the merge's afterwards.
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

  it("offers to REMOVE the preview while previewing, not to preview again", () => {
    // "Put the template back" was the old design's language, and the old
    // design is on this project's rejected list: it wrote the row onto the
    // template through an API that re-authors text. Nothing is taken from the
    // template now, so nothing is put back — slides are deleted.
    const label = primary({ ...ready, previewing: true }, "preview").label;
    expect(label).toBe("Remove the preview");
    expect(label).not.toContain("template");
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

describe("reading the two slide-number boxes", () => {
  it("says nothing at all while a box is still empty", () => {
    // A form that turns red on the first keystroke is wrong more often than
    // the user is. The boxes are filled one at a time, so "4" and "" is a
    // half-typed entry, not a mistake.
    for (const draft of [
      { from: "", to: "" },
      { from: "4", to: "" },
      { from: "", to: "6" },
    ]) {
      const read = readBlockDraft(draft);
      expect(read.why, JSON.stringify(draft)).toBeNull();
      expect(read.block, JSON.stringify(draft)).toBeNull();
    }
  });

  it("reads two numbers into a block", () => {
    expect(readBlockDraft({ from: "4", to: "6" }).block).toEqual({ from: 4, to: 6 });
  });

  it("ignores the whitespace a paste brings with it", () => {
    expect(readBlockDraft({ from: " 4 ", to: "\t6" }).block).toEqual({ from: 4, to: 6 });
  });

  it("refuses a block that ends before it starts, naming both slides", () => {
    const read = readBlockDraft({ from: "6", to: "4" });
    expect(read.block).toBeNull();
    expect(read.why).toContain("6");
    expect(read.why).toContain("4");
  });

  it("refuses slide 0, because the rail starts at 1", () => {
    // The REFUSAL, not only the sentence. This asserted `why` contained "1"
    // and nothing else — satisfied by the literal 1 in "numbered from 1", so
    // an implementation returning a block AND a complaint passed the whole
    // suite, and the user then spent a real host round trip on slide 0.
    const read = readBlockDraft({ from: "0", to: "3" });
    expect(read.block).toBeNull();
    expect(read.why).toContain("0");
  });

  it("refuses a fraction rather than rounding one", () => {
    // Rounding picks a slide the user did not name, and the merge then clones
    // it perfectly.
    expect(readBlockDraft({ from: "1.5", to: "3" }).block).toBeNull();
    expect(readBlockDraft({ from: "1", to: "abc" }).block).toBeNull();
  });

  it("names what the USER typed in every refusal, not just the rule", () => {
    // "Slides are numbered from 1." is a true sentence that says nothing about
    // the boxes in front of them — and the manual promised numbers for all
    // four cases while two of them carried none.
    expect(readBlockDraft({ from: "0", to: "3" }).why).toContain("0");
    expect(readBlockDraft({ from: "1.5", to: "3" }).why).toContain("1.5");
    expect(readBlockDraft({ from: "6", to: "4" }).why).toContain("6");
    expect(readBlockDraft({ from: "6", to: "4" }).why).toContain("4");
  });

  it("WARNS about a block past the end of the deck, and still lets it through", () => {
    // Advice, not a refusal. `deckSize` is counted once when the pane opens, so
    // a user who adds slides and comes back would otherwise be told their block
    // does not exist — in a sentence stating a deck size that is no longer
    // true, with no way to correct it short of reopening the pane. `blockIds`
    // checks it a moment later against ids the host listed just now.
    const read = readBlockDraft({ from: "4", to: "9" }, 6);
    expect(read.block, "the user may press past it").toEqual({ from: 4, to: 9 });
    expect(read.why).toContain("9");
    expect(read.why).toContain("6");
    // And nothing is said at all before the deck has answered.
    expect(readBlockDraft({ from: "4", to: "9" }).why).toBeNull();
  });

  it("says slide, singular, when the deck holds exactly one", () => {
    expect(readBlockDraft({ from: "1", to: "2" }, 1).why).toContain("1 slide");
    expect(readBlockDraft({ from: "1", to: "2" }, 1).why).not.toContain("1 slides");
  });
});

describe("which block the pane acts on", () => {
  it("prefers what the boxes say over what was committed", () => {
    // A button reading "Use slides 4 to 6" while the boxes hold 7 and 9 is a
    // button that does something other than what it says.
    const state: PaneState = { ...ready, draft: { from: "7", to: "9" } };
    expect(chosenBlock(state)).toEqual({ from: 7, to: 9 });
    expect(primary(state, "template").label).toBe("Use slides 7 to 9");
  });

  it("keeps the committed block while a box is being retyped", () => {
    // Mid-edit the draft reads nothing. Dropping the block there would blank
    // the heading and the button on every keystroke.
    expect(chosenBlock({ ...ready, draft: { from: "7", to: "" } })).toEqual({ from: 4, to: 6 });
  });

  it("is undefined when there is neither", () => {
    expect(chosenBlock(EMPTY)).toBeUndefined();
  });

  it("unblocks the later steps off the draft alone", () => {
    const typed: PaneState = { fields: [], previewing: false, draft: { from: "2", to: "4" } };
    expect(blockedReason(typed, "fields")).toBeNull();
  });
});

describe("reading what was pasted", () => {
  it("says nothing about an empty box", () => {
    expect(readPastedTable("")).toEqual({ records: null, columns: [], rows: 0, why: null });
    expect(readPastedTable("   \n ").why).toBeNull();
  });

  it("reads a range pasted out of Excel, which arrives tab-separated", () => {
    const read = readPastedTable("First\tLast\nAda\tLovelace\nGrace\tHopper");
    expect(read.columns).toEqual(["First", "Last"]);
    expect(read.rows).toBe(2);
    expect(read.records?.rows[0]).toEqual({ First: "Ada", Last: "Lovelace" });
  });

  it("reads a comma-separated paste too", () => {
    expect(readPastedTable("First,Last\nAda,Lovelace").columns).toEqual(["First", "Last"]);
  });

  it("counts the rows the merge will actually run, not the lines pasted", () => {
    // A trailing newline and a blank line in the middle are both what a copy
    // out of a spreadsheet brings with it, and neither is a row.
    const read = readPastedTable("Name\nAda\n\nGrace\n");
    expect(read.rows).toBe(2);
    expect(read.records?.rows).toHaveLength(2);
  });

  it("refuses a header row with nothing under it", () => {
    // The button would otherwise say "Add 0 slides" and the user would have no
    // idea which half was wrong.
    expect(readPastedTable("First\tLast").why).toContain("header");
    expect(readPastedTable("First\tLast").records).toBeNull();
  });

  it("is the ONE parse: the columns shown are the columns the merge binds", () => {
    // Two parses — one for the labels, one for the merge — is two that can
    // disagree, and the one that disagrees is the one nobody sees.
    const read = readPastedTable("A\tA\nx\ty");
    expect(read.columns).toEqual(read.records?.columns.map((c) => c.name));
  });
});

describe("moving between steps", () => {
  it("gives the next step in order", () => {
    expect(nextStep("template")).toBe("fields");
    expect(nextStep("fields")).toBe("preview");
    expect(nextStep("preview")).toBe("merge");
  });

  it("has nowhere to go after the last step", () => {
    expect(nextStep("merge")).toBeNull();
  });

  it("does NOT send an unknown step back to the beginning", () => {
    // `order[order.indexOf(from) + 1]` answers order[0] for anything that is
    // not a step, so a stray data-action put the user on step 1 with their
    // block and their data still in state — a wizard that resets itself and
    // looks like it lost the lot.
    expect(nextStep("nonsense" as StepId)).toBeNull();
  });
});

describe("what the primary says once there is data", () => {
  it("stops saying Attach data after the data is attached", () => {
    // A button that still says "Attach data" reads as a step that did not take.
    expect(primary(ready, "fields").label).toBe("Use 240 rows");
    expect(primary({ ...ready, rows: 1 }, "fields").label).toBe("Use 1 row");
  });

  it("cannot be pressed with nothing pasted", () => {
    const noData: PaneState = { block: { from: 4, to: 6 }, fields: [], previewing: false };
    expect(primary(noData, "fields")).toEqual({ label: "Attach data", enabled: false });
  });

  it("offers the preview, now that pressing it shows one", () => {
    expect(primary(ready, "preview").label).toBe("Preview the first row");
    expect(primary(ready, "preview").enabled).toBe(true);
    // And is not offerable before there is a row to show.
    expect(primary(EMPTY, "preview").enabled).toBe(false);
  });
});

describe("which row a preview shows", () => {
  const records = {
    columns: [{ name: "First", type: "text" as const }],
    rows: [{ First: "Ada" }, { First: "Grace" }, { First: "Katherine" }],
  };

  it("is the first, and only the first", () => {
    const one = firstRowOnly(records);
    expect(one.rows).toEqual([{ First: "Ada" }]);
  });

  it("keeps every column, so the merge binds what it would bind", () => {
    // A preview that dropped a column would report unmatched placeholders the
    // real merge does not have — a preview of something nobody is going to get.
    expect(firstRowOnly(records).columns).toEqual(records.columns);
  });

  it("answers an empty set with an empty set rather than a row of nothing", () => {
    expect(firstRowOnly({ columns: records.columns, rows: [] }).rows).toEqual([]);
  });
});

describe("taking rows out of the merge", () => {
  const records = {
    columns: [
      { name: "Name", type: "text" as const },
      { name: "City", type: "text" as const },
    ],
    rows: [
      { Name: "Ada", City: "London" },
      { Name: "Grace", City: "New York" },
      { Name: "", City: "Aarhus" },
      { Name: "Katherine", City: "Hampton" },
    ],
  };
  const withData: PaneState = { ...ready, records, rows: 4 };

  it("labels a row by its first column", () => {
    expect(rowLabel(records, 0)).toBe("Ada");
  });

  it("falls back to the position when that cell is empty", () => {
    // The row still needs something to click.
    expect(rowLabel(records, 2)).toBe("Row 3");
  });

  it("searches EVERY column, not just the labelled one", () => {
    // Someone looking for Aarhus is looking for the row with Aarhus in it, and
    // whether that is the column the label came from is not something they
    // should have to know.
    expect(visibleRows(records, "aarhus")).toEqual([2]);
    expect(visibleRows(records, "ada")).toEqual([0]);
  });

  it("matches case-insensitively, and everything on an empty query", () => {
    expect(visibleRows(records, "LONDON")).toEqual([0]);
    expect(visibleRows(records, "")).toEqual([0, 1, 2, 3]);
    expect(visibleRows(records, "   ")).toEqual([0, 1, 2, 3]);
  });

  it("merges everything when nothing has been touched", () => {
    // The default has to be "all", which is why the state holds the EXCLUDED
    // rows: an included-list would default to empty and merge nothing.
    expect(includedCount(withData)).toBe(4);
    expect(includedRecords(withData)?.rows).toHaveLength(4);
  });

  it("leaves out the rows that were taken out, keeping order and columns", () => {
    const some: PaneState = { ...withData, excluded: [1, 3] };
    expect(includedCount(some)).toBe(2);
    const left = includedRecords(some);
    expect(left?.rows.map((r) => r.Name)).toEqual(["Ada", ""]);
    // A filter removes rows and nothing else.
    expect(left?.columns).toEqual(records.columns);
  });

  it("counts a duplicate exclusion once", () => {
    // A repeated index would otherwise subtract the same row twice and report
    // fewer rows than will merge.
    expect(includedCount({ ...withData, excluded: [1, 1, 1] })).toBe(3);
  });

  it("ignores an exclusion that is not a row", () => {
    expect(includedCount({ ...withData, excluded: [-1, 99] })).toBe(4);
  });

  it("counts from `rows` when the state has a count but no records", () => {
    // A state can carry the count without the data — the pane knows how many
    // rows it has before it needs them. Reading only `records` made this
    // answer ZERO for every such state, which blocked the merge and emptied
    // the button's number.
    expect(includedCount({ ...ready, rows: 240 })).toBe(240);
    expect(includedCount({ ...ready, rows: 240, excluded: [0, 1] })).toBe(238);
  });

  it("puts the INCLUDED count on the button, not the pasted one", () => {
    const some: PaneState = { ...withData, excluded: [1, 3] };
    // 2 rows x 3 slides.
    expect(primary(some, "merge").label).toBe("Add 6 slides");
  });

  it("blocks the merge when every row is unticked, and says which it is", () => {
    // Not the same as having no data at all: this was deliberate, so the
    // sentence has to name what the user did rather than say "attach data".
    const none: PaneState = { ...withData, excluded: [0, 1, 2, 3] };
    expect(blockedReason(none, "merge")).toContain("unticked");
    expect(primary(none, "merge").enabled).toBe(false);
  });

  it("previews the first row that will MERGE, not the first pasted", () => {
    // A preview of a row the user has unticked is a preview of something
    // nobody is going to get.
    const some: PaneState = { ...withData, excluded: [0] };
    expect(firstIncludedRow(some)?.rows).toEqual([{ Name: "Grace", City: "New York" }]);
  });

  it("has no row to preview when they are all out", () => {
    expect(firstIncludedRow({ ...withData, excluded: [0, 1, 2, 3] })).toBeUndefined();
  });
});
