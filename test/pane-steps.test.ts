import { describe, expect, it } from "vitest";
import {
  announcement,
  blockSlides,
  conditionFor,
  danglingConditions,
  withCondition,
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
  slidesToAdd,
  readBlockDraft,
  readPastedTable,
  rowLabel,
  fieldToken,
  imageColumns,
  plannedSlides,
  pictureColumns,
  imageTally,
  imageNameClashes,
  clashingPicturesNote,
  imagesWanted,
  slidesPerRecord,
  statusOf,
  caution,
  skippedRows,
  unmatchedFields,
  visibleRows,
} from "../src/pane/steps.js";
import type { Block, PaneState, StepId } from "../src/pane/steps.js";
import { toRecordSet, type RecordSet } from "../src/core/data/recordset.js";
import { buildPlan, slideCount } from "../src/core/merge/plan.js";
import { blockMoved } from "../src/pane/transitions.js";
import { fieldsInText } from "../src/core/merge/text.js";
import { imageMode } from "../src/core/merge/images.js";

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
    for (const step of ["data", "fields", "preview", "merge"] as const) {
      expect(blockedReason(EMPTY, step), step).toContain("repeat");
    }
  });

  it("blocks fields, preview and merge until data is attached", () => {
    // A field IS a column name, so the fields step has nothing to offer until
    // the data is attached. That is the whole reason the order is template,
    // data, fields rather than template, fields, data.
    const noData: PaneState = { block: { from: 4, to: 6 }, fields: [], previewing: false };
    expect(blockedReason(noData, "data")).toBeNull();
    expect(blockedReason(noData, "fields")).toContain("data");
    expect(blockedReason(noData, "preview")).toContain("data");
    expect(blockedReason(noData, "merge")).toContain("data");
  });

  it("blocks preview and merge until something is on the slides", () => {
    // The engine refuses a block with no placeholders too, and must: N
    // identical copies is never what anybody meant and is expensive to undo.
    // Said here so the refusal costs no host call and no insert.
    const noFields: PaneState = {
      block: { from: 4, to: 6 },
      fields: [],
      previewing: false,
      columns: ["First"],
      rows: 3,
    };
    expect(blockedReason(noFields, "fields"), "the step that fixes it").toBeNull();
    expect(blockedReason(noFields, "preview")).toContain("no fields yet");
    expect(blockedReason(noFields, "merge")).toContain("no fields yet");
    // Names the slides, because the user has to go and look at them.
    expect(blockedReason(noFields, "merge")).toContain("Slides 4 to 6");
  });

  it("NAMES the placeholders that have no column", () => {
    // A count alone sends the user back through every slide looking for it.
    const missing: PaneState = { ...ready, fields: ["First", "Nickname", "Badge"] };
    const said = caution(missing, "merge") ?? "";
    expect(said).toContain("Nickname");
    expect(said).toContain("Badge");
    expect(said).not.toContain("First");
  });

  it("does not REFUSE the merge over them", () => {
    // It did until 2026-08-29, and three other parts of this project disagreed:
    // the engine leaves such a placeholder on the slide on purpose, the preview
    // step ran the ordinary merge with one, and docs/MANUAL.md promises it. The
    // asymmetry was the tell — a field whose column was missing was refused, a
    // field whose PICTURE was missing was allowed and documented, and both end
    // with a placeholder on the slide.
    const missing: PaneState = { ...ready, fields: ["First", "Nickname"] };
    expect(blockedReason(missing, "merge")).toBeNull();
    expect(primary(missing, "merge").enabled).toBe(true);
  });

  it("says what will happen rather than what to fix", () => {
    // "Rename the column or the placeholder" reads as a demand to correct a
    // mistake. Staying on the slide is the documented behaviour, so the
    // sentence says that instead.
    const one: PaneState = { ...ready, fields: ["First", "Nickname"] };
    expect(caution(one, "merge")).toContain("It will stay on the slides as written");
    const two: PaneState = { ...ready, fields: ["First", "Nickname", "Badge"] };
    expect(caution(two, "merge")).toContain("They will stay on the slides as written");
  });

  it("is silent on every other step, and when everything matches", () => {
    expect(caution(ready, "merge"), "nothing is unmatched").toBeNull();
    const missing: PaneState = { ...ready, fields: ["First", "Nickname"] };
    for (const step of ["template", "data", "fields", "preview"] as const) {
      expect(caution(missing, step), step).toBeNull();
    }
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
    expect(label).toBe("On to the merge");
    expect(label).not.toContain("template");
  });

  it("carries ON to the merge while previewing, rather than only clearing up", () => {
    // The primary used to read "Remove the preview" and stop there, which left
    // step 4 holding "Back to fields" and that button and nothing else — the
    // word "merge" appeared nowhere on the screen, and the route on was to work
    // out that clearing the preview was the route. Removing is still offered,
    // on the card beside the slides it names; the primary is the way onward,
    // as it is on every other step.
    const action = primary({ ...ready, previewing: true }, "preview");
    expect(action.label).toBe("On to the merge");
    expect(action.advances, "the primary has to be the door here, or the step has none").toBe("merge");
  });

  it("says where every primary that moves the wizard lands, pressable or not", () => {
    // `advances` is where the door GOES; `enabled` is whether it is open yet.
    // Step 1 opens disabled and is still the way to step 2 — conflating the two
    // is exactly what would make a walk call that opening screen a dead end.
    expect(primary(EMPTY, "template")).toMatchObject({ enabled: false, advances: "data" });
    expect(primary(ready, "template")).toMatchObject({ enabled: true, advances: "data" });
    // A primary that ACTS rather than advances says nothing, and the forward
    // link beside it is the door instead.
    expect(primary(ready, "preview").advances).toBeUndefined();
    // The last step has nowhere to advance to.
    expect(primary(ready, "merge").advances).toBeUndefined();
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
    expect(blockedReason(typed, "data")).toBeNull();
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
    expect(nextStep("template")).toBe("data");
    expect(nextStep("data")).toBe("fields");
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
    expect(primary(ready, "data").label).toBe("Use 240 rows");
    expect(primary({ ...ready, rows: 1 }, "data").label).toBe("Use 1 row");
  });

  it("cannot be pressed with nothing pasted", () => {
    const noData: PaneState = { block: { from: 4, to: 6 }, fields: [], previewing: false };
    expect(primary(noData, "data")).toMatchObject({ label: "Attach data", enabled: false });
  });

  it("names what the fields step will do, and what it already knows", () => {
    // One press, one job: read the slides again and go on. The user has just
    // been putting `{{Column}}` onto them and nothing tells the pane that
    // happened — there is no document-changed event for slide text.
    const nothingYet: PaneState = { block: { from: 4, to: 6 }, fields: [], previewing: false, rows: 3 };
    expect(primary(nothingYet, "fields")).toMatchObject({ label: "Check the slides for fields", enabled: true });
    expect(primary(ready, "fields").label).toBe("Use 2 fields");
    expect(primary({ ...ready, fields: ["First"] }, "fields").label).toBe("Use 1 field");
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

describe("conditional slides", () => {
  const block = { from: 4, to: 6 };
  const base: PaneState = {
    block,
    fields: ["First"],
    previewing: false,
    columns: ["First", "Last", "Email"],
    rows: 2,
  };

  it("names the slides in the block the way the user sees them", () => {
    // The key the engine reads back: `prepareBlock` looks up
    // `conditions[req.from + i]`, so pane and engine agree on "slide 5" without
    // either converting. A block of one slide is one number, not none.
    expect(blockSlides(base)).toEqual([4, 5, 6]);
    expect(blockSlides({ ...base, block: { from: 2, to: 2 } })).toEqual([2]);
    expect(blockSlides({ fields: [], previewing: false })).toEqual([]);
  });

  it("stores a choice and reads it back", () => {
    const conditions = withCondition(undefined, 5, "Email");
    expect(conditions).toEqual({ 5: "Email" });
    expect(conditionFor({ ...base, conditions }, 5)).toBe("Email");
    expect(conditionFor({ ...base, conditions }, 4)).toBe("");
  });

  it("deletes rather than storing a blank, and carries no empty object", () => {
    /**
     * `prepareBlock` tests the value for truthiness, so a stored "" behaves as
     * "always" — right by accident, and still counted as a condition by
     * anything reading keys, which is what the summary line does.
     *
     * `undefined` rather than `{}` for the same reason one meaning gets one
     * spelling: two readers comparing "has conditions" would otherwise
     * disagree.
     */
    const one = withCondition(undefined, 5, "Email");
    expect(withCondition(one, 5, "")).toBeUndefined();
    expect(withCondition({ 4: "Last", 5: "Email" }, 5, "")).toEqual({ 4: "Last" });
  });

  it("does not mutate the conditions it was given", () => {
    // The pane rebuilds its state from the old one on every change; a mutated
    // record would be shared with the state a host call is about to answer
    // into.
    const before = { 5: "Email" };
    withCondition(before, 4, "Last");
    expect(before).toEqual({ 5: "Email" });
  });

  it("reports a condition naming a column the data does not have", () => {
    // Reachable without anyone typing a name: choose a column, then paste
    // different data. The engine emits the slide anyway rather than hiding an
    // authoring mistake, so the pane says so while it is still free to fix.
    const state = { ...base, conditions: { 4: "Renewal", 6: "Email" } };
    expect(danglingConditions(state)).toEqual(["Renewal"]);
    expect(danglingConditions({ ...base, conditions: { 6: "Email" } })).toEqual([]);
  });

  it("says nothing about conditions on slides outside the block", () => {
    // A stale key cannot be reported as a problem the user can act on: there is
    // no control for slide 9 while the block is 4 to 6, so naming it would send
    // them looking for one.
    expect(danglingConditions({ ...base, conditions: { 9: "Renewal" } })).toEqual([]);
  });

  it("cannot report anything before data is attached", () => {
    // Every column is unknown with no data, which would flag every condition.
    expect(danglingConditions({ ...base, columns: undefined, conditions: { 4: "Renewal" } })).toEqual([]);
  });
});

/**
 * The pane fields a paste produces, spread the way `main.ts` spreads them.
 *
 * Not `...readPastedTable(text)`: that carries a `records: null` and a `why`,
 * neither of which is a `PaneState` field. A fixture built from the parser
 * rather than described by hand, so it cannot disagree with what the pane holds.
 */
function attached(text: string): Pick<PaneState, "paste" | "records" | "columns" | "rows"> {
  const read = readPastedTable(text);
  return { paste: text, records: read.records ?? undefined, columns: read.columns, rows: read.rows };
}

describe("which columns name pictures", () => {
  const state = (text: string): PaneState => ({ ...EMPTY, ...attached(text) });

  it("reads the image columns off the parse rather than off the header", () => {
    expect(imageColumns(state("Name,Photo\nAda,ada.png\nGrace,grace.jpg"))).toEqual(["Photo"]);
  });

  it("does not call a column of ordinary words an image column", () => {
    expect(imageColumns(state("Name,City\nAda,London"))).toEqual([]);
  });

  it("names every picture the data asks for, once each", () => {
    const s = state("Name,Photo\nAda,ada.png\nGrace,ada.png\nAlan,alan.jpg");
    expect(imagesWanted(s).sort()).toEqual(["ada.png", "alan.jpg"]);
  });
});

describe("what the picked files cover", () => {
  const data = (): PaneState => ({
    ...EMPTY,
    ...attached("Name,Photo\nAda,ada.png\nGrace,grace.JPG\nAlan,alan.png"),
  });

  const files = (...names: string[]): Map<string, Uint8Array> => new Map(names.map((n) => [n, new Uint8Array([1])]));

  it("counts nothing matched before any file is picked", () => {
    expect(imageTally(data())).toMatchObject({ wanted: 3, matched: 0, spare: [] });
  });

  it("matches by base name, so a folder path in the cell is not a miss", () => {
    const s: PaneState = {
      ...EMPTY,
      ...attached("Name,Photo\nAda,Photos\\ada.png"),
      images: files("ada.png"),
    };
    expect(imageTally(s)).toMatchObject({ wanted: 1, matched: 1, missing: [] });
  });

  it("matches case-insensitively, the way the merge matches", () => {
    const s = { ...data(), images: files("ADA.PNG", "grace.jpg", "alan.png") };
    expect(imageTally(s)).toMatchObject({ matched: 3, missing: [] });
  });

  it("names the pictures it has not got, rather than counting them", () => {
    const s = { ...data(), images: files("ada.png") };
    expect(imageTally(s).missing.sort()).toEqual(["alan.png", "grace.JPG"]);
  });

  it("counts a file no row refers to as spare, not as a problem", () => {
    const s = { ...data(), images: files("ada.png", "grace.jpg", "alan.png", "logo.png") };
    expect(imageTally(s)).toMatchObject({ matched: 3, missing: [], spare: ["logo.png"] });
  });

  /**
   * The hole the two rules above leave between them.
   *
   * Cells may carry folders and matching is by base name — both deliberate,
   * because a browser's file picker hands back `File.name` and that has no
   * path. Two cells that differ ONLY by the folder are therefore two pictures
   * to the author and one name to everything else: `imageTally` calls both
   * matched, the merge fills both from the same file, and one slide carries
   * the wrong picture with nothing having said so.
   */
  it("says when two picture names differ only by their folder", () => {
    const s: PaneState = {
      ...EMPTY,
      ...attached("Name,Photo\nAda,regions/eu/logo.png\nBo,regions/us/logo.png"),
      images: files("logo.png"),
    };
    // The tally is not wrong — both DO resolve to a file that was attached.
    expect(imageTally(s)).toMatchObject({ wanted: 2, matched: 2, missing: [] });
    expect(imageNameClashes(s)).toEqual([["regions/eu/logo.png", "regions/us/logo.png"]]);
    const note = clashingPicturesNote(s) ?? "";
    expect(note).toContain("regions/eu/logo.png");
    expect(note).toContain("regions/us/logo.png");
    expect(note).toContain("same file name in different folders");
  });

  it("says nothing when the names are genuinely different", () => {
    expect(imageNameClashes(data())).toEqual([]);
    expect(clashingPicturesNote(data())).toBeNull();
  });

  it("counts the rest of the clashes rather than listing every one", () => {
    const s: PaneState = {
      ...EMPTY,
      ...attached("Name,Photo\nA,a/x.png\nB,b/x.png\nC,a/y.png\nD,b/y.png"),
    };
    expect(imageNameClashes(s)).toHaveLength(2);
    expect(clashingPicturesNote(s)).toContain("One more name clashes");
  });

  it("does not call the same cell in two rows a clash", () => {
    // `ada.png` twice is one picture wanted twice, which is the ordinary case.
    const s: PaneState = { ...EMPTY, ...attached("Name,Photo\nAda,ada.png\nBo,ada.png") };
    expect(imageNameClashes(s)).toEqual([]);
  });
});

describe("the token an image column is written as", () => {
  it("asks for a picture, because the engine decides that from the format", () => {
    expect(fieldToken("Photo", "image")).toBe("{{Photo|image}}");
  });

  it("leaves every other column alone", () => {
    expect(fieldToken("Photo")).toBe("{{Photo}}");
    expect(fieldToken("City", "text")).toBe("{{City}}");
  });

  it("is a token the engine's own reader accepts, name and format both", () => {
    const hits = fieldsInText(fieldToken("Photo", "image"));
    expect(hits).toHaveLength(1);
    expect(hits[0]?.name).toBe("Photo");
    expect(imageMode(hits[0]?.format)).toBe("cover");
  });
});

describe("a blocked step never offers a pressable button", () => {
  /**
   * `blockedReason` promises that every sentence "names the thing the user has
   * to do". A step that shows one of those sentences above a button that WORKS
   * is naming the wrong thing at the one moment the user has an obvious right
   * one to do.
   *
   * Swept rather than listed. The state that broke it was reachable by a route
   * nobody would write a case for: preview a row, go back to step 1, and type
   * in a slide-number box. That clears the committed block — deliberately, it
   * is stale — and does not end the preview, so step 4 read "Choose the slides
   * that repeat first." directly above a working "Remove the preview".
   *
   * Never stuck, to be clear: navigation is ungated on purpose and the button
   * always worked. It was the sentence that was wrong.
   */
  const STATES: PaneState[] = (() => {
    const out: PaneState[] = [];
    for (const block of [undefined, { from: 1, to: 2 }])
      for (const rows of [undefined, 0, 2])
        for (const fields of [[], ["Name"], ["Name", "Nickname"]])
          for (const previewing of [false, true])
            for (const running of [undefined, "inspect", "merge", "preview"] as const)
              for (const added of [undefined, 6])
                for (const excluded of [undefined, [0, 1]])
                  out.push({
                    ...(block ? { block } : {}),
                    fields,
                    ...(rows === undefined ? {} : { rows, columns: ["Name", "Region"] }),
                    previewing,
                    ...(running ? { running } : {}),
                    ...(added === undefined ? {} : { added }),
                    ...(excluded ? { excluded } : {}),
                  });
    return out;
  })();

  it("across every combination of the things a step gates on", () => {
    const violations: string[] = [];
    for (const state of STATES) {
      for (const step of STEPS) {
        const why = blockedReason(state, step);
        const button = primary(state, step);
        if (why !== null && button.enabled) violations.push(`${step}: "${why}" over a live "${button.label}"`);
      }
    }
    expect([...new Set(violations)]).toEqual([]);
    expect(STATES.length, "the sweep stopped covering anything").toBeGreaterThan(500);
  });

  it("and the route that broke it reads the right way round", () => {
    // The block goes, the preview stays. Step 4 has one thing to say now.
    const previewing: PaneState = { fields: ["Name"], rows: 2, columns: ["Name"], previewing: true };
    expect(blockedReason(previewing, "preview")).toBeNull();
    expect(primary(previewing, "preview")).toMatchObject({ label: "On to the merge", enabled: true });
  });

  it("without making the step reachable when no preview is on the slides", () => {
    // The other half: with nothing previewing, the three sentences still stand.
    const bare: PaneState = { fields: [], previewing: false };
    expect(blockedReason(bare, "preview")).toBe("Choose the slides that repeat first.");
    expect(primary(bare, "preview").enabled).toBe(false);
  });
});

describe("a picture column one stray cell kept out of the type", () => {
  /**
   * `detectType` is all-or-nothing on purpose: one cell reading `n/a` in a
   * column of file names makes the whole column text, so a column of `.svg`
   * names is not offered as pictures and then failed one row at a time.
   *
   * The pane decided what a picture was from that alone. The ENGINE decides
   * from the FIELD's format — `{{Photo|image}}` is documented in
   * `docs/MANUAL.md` as the way to ask for one, and `placeImages` obeys it
   * whatever the column's type.
   *
   * So an author who wrote the format by hand on such a column got a pane with
   * NO picker — it is shown only when `imagesWanted` is non-empty — and
   * therefore no way to attach the files, and a merge that left every picture
   * placeholder standing. The insert button made it worse by writing
   * `{{Photo}}`, which merges the file name as text.
   *
   * `imageFieldsIn` has answered "which fields ask for a picture" since it was
   * written, and nothing in the product called it.
   */
  const stray = toRecordSet([["Photo"], ["ada.png"], ["n/a"]]);

  const withField = (imageFields: string[]): PaneState => ({
    fields: ["Photo"],
    imageFields,
    previewing: false,
    records: stray,
    columns: ["Photo"],
    rows: stray.rows.length,
  });

  it("is not an image column, and that part is deliberate", () => {
    expect(stray.columns[0]?.type).toBe("text");
    expect(imageColumns(withField(["Photo"]))).toEqual([]);
  });

  it("is still a picture column when the author asked for one", () => {
    expect(pictureColumns(withField(["Photo"]))).toEqual(["Photo"]);
    expect(imagesWanted(withField(["Photo"])).length, "the picker never appears").toBeGreaterThan(0);
  });

  it("and the insert button writes the format the engine acts on", () => {
    const state = withField(["Photo"]);
    expect(fieldToken("Photo", pictureColumns(state).includes("Photo") ? "image" : undefined)).toBe("{{Photo|image}}");
  });

  it("without inventing one the author did not ask for", () => {
    // The other half: no image field, no picker. A text column stays text.
    expect(pictureColumns(withField([]))).toEqual([]);
    expect(imagesWanted(withField([]))).toEqual([]);
  });

  it("and ignores an image field naming a column the data does not have", () => {
    // That is an unmatched field, which the fields step already reports.
    expect(pictureColumns(withField(["Nickname"]))).toEqual([]);
  });
});

describe("the number above the merge button", () => {
  /**
   * It said "9 slides added after slide 10, leaving 19 slides in the deck" and
   * the plan built eight.
   *
   * The count was slides-per-record times rows, which knows nothing about
   * CONDITIONS: a block with one conditional slide produces fewer slides for
   * every row the condition leaves out. So the sentence a user reads to decide
   * whether to press was over by one per skipped slide, and the deck size it
   * predicted was wrong with it.
   *
   * `plannedSlides` counts with `slideApplies` — the rule `buildPlan` itself
   * applies — so the promise and the plan cannot answer differently. This test
   * asserts that agreement rather than the number, because the number is only
   * right for as long as the two rules are one.
   */
  const records = toRecordSet([
    ["Name", "Renewal"],
    ["Ada", "yes"],
    ["Bo", "no"],
    ["Cy", "yes"],
    ["Di", "no"],
  ]);

  const state: PaneState = {
    block: { from: 3, to: 5 },
    fields: ["Name"],
    previewing: false,
    records,
    columns: ["Name", "Renewal"],
    rows: records.rows.length,
    deckSize: 10,
    excluded: [3],
    conditions: { 4: "Renewal" },
  };

  /** The same block, as the engine takes it. */
  const engineBlock = {
    id: "b",
    slides: [
      { path: "s3.xml", seq: 1 },
      { path: "s4.xml", seq: 2, condition: "Renewal" },
      { path: "s5.xml", seq: 3 },
    ],
  };

  it("is what the plan will actually build", () => {
    const chosen = includedRecords(state) as RecordSet;
    const plan = buildPlan(engineBlock, chosen, { runId: "r" });
    expect(plannedSlides(state)).toBe(slideCount(plan));
    // And it is not the product that used to be shown, or the test would pass
    // against the bug.
    expect(plannedSlides(state)).not.toBe(slidesPerRecord(state.block as Block) * includedCount(state));
  });

  /**
   * The same agreement under `onEmpty: "skip"`, and the case that decides the
   * whole design.
   *
   * `buildPlan` drops a record when a field on one of the slides THAT RECORD
   * ACTUALLY GETS is blank — the slides its own conditions leave in. A pane
   * counting from the flat field list would drop a row over a blank field on a
   * slide the row's condition had already removed, and put a number on the
   * button the plan does not produce. That is why `slideFields` is carried out
   * of the template read at all.
   */
  describe("under 'leave the whole row out'", () => {
    const blanks = toRecordSet([
      ["Name", "Renewal", "Note"],
      // Gets slide 4 (Renewal yes) and its Note is blank: dropped.
      ["Ada", "yes", ""],
      // Its Note is blank too, but Renewal is no, so slide 4 — the only slide
      // carrying {{Note}} — is not one of Bo's. NOT dropped.
      ["Bo", "no", ""],
      ["Cy", "yes", "kept"],
    ]);

    /** Slide 4 is the conditional one, and the only one holding {{Note}}. */
    const perSlide = [["Name"], ["Name", "Note"], ["Name"]];

    const skipping: PaneState = {
      block: { from: 3, to: 5 },
      fields: ["Name", "Note"],
      slideFields: perSlide,
      previewing: false,
      records: blanks,
      columns: ["Name", "Renewal", "Note"],
      rows: blanks.rows.length,
      deckSize: 10,
      conditions: { 4: "Renewal" },
      onEmpty: "skip",
    };

    const engineSkipping = {
      id: "b",
      slides: [
        { path: "s3.xml", seq: 1, fields: perSlide[0] },
        { path: "s4.xml", seq: 2, condition: "Renewal", fields: perSlide[1] },
        { path: "s5.xml", seq: 3, fields: perSlide[2] },
      ],
    };

    it("drops the rows the plan drops, and no others", () => {
      const plan = buildPlan(engineSkipping, blanks, { runId: "r", onEmpty: "skip" });
      expect(plan.skippedRecords, "only Ada, whose own slides carry the blank").toEqual([0]);
      expect(skippedRows(skipping)).toEqual(plan.skippedRecords);
    });

    it("counts what the plan will actually build", () => {
      const plan = buildPlan(engineSkipping, blanks, { runId: "r", onEmpty: "skip" });
      expect(plannedSlides(skipping)).toBe(slideCount(plan));
      // Not the product, or this passes against a pane that ignores the policy.
      expect(plannedSlides(skipping)).not.toBe(slidesPerRecord(skipping.block as Block) * includedCount(skipping));
    });

    it("counts nothing dropped under the other two answers", () => {
      expect(skippedRows({ ...skipping, onEmpty: "blank" })).toEqual([]);
      expect(skippedRows({ ...skipping, onEmpty: "keep" })).toEqual([]);
      expect(skippedRows({ ...skipping, onEmpty: undefined })).toEqual([]);
    });

    it("counts nothing before the slides have been read", () => {
      // `slideFields` arrives with the template read. Guessing from the flat
      // list in the meantime would be the second rule this avoids.
      expect(skippedRows({ ...skipping, slideFields: undefined })).toEqual([]);
    });

    it("names the count above the button, before the press", () => {
      expect(caution(skipping, "merge") ?? "").toContain("1 of 3 rows will be left out");
    });

    it("previews a row that will actually merge", () => {
      // The rule `firstIncludedRow` already followed for an unticked row: a
      // preview of a row nobody gets is a preview of nothing. Ada is dropped,
      // so the preview is Bo.
      expect(firstIncludedRow(skipping)?.rows[0]?.Name).toBe("Bo");
    });

    it("ignores a blank column that is on no slide", () => {
      // A spreadsheet routinely carries columns a template does not use, and
      // half of them are empty. Dropping a row for one of those would make the
      // control useless on real data — and it is the pane, not the engine,
      // that puts the number on the button.
      const spare = toRecordSet([
        ["Name", "Renewal", "Note", "Comment"],
        ["Ada", "no", "kept", ""],
        ["Bo", "no", "kept", ""],
      ]);
      const withSpare = {
        ...skipping,
        records: spare,
        rows: spare.rows.length,
        columns: [...spare.columns.map((c) => c.name)],
      };
      expect(skippedRows(withSpare)).toEqual([]);
      expect(caution(withSpare, "merge") ?? "").not.toContain("will be left out");
      // And the count is what it would be without the policy at all — which is
      // the claim, and not the product: a condition is live on this fixture,
      // so slides-per-record times rows is the wrong comparison and asserting
      // it fails on arithmetic that has nothing to do with blank cells.
      expect(plannedSlides(withSpare)).toBe(plannedSlides({ ...withSpare, onEmpty: "blank" }));
    });

    it("ignores a field the data has no column for", () => {
      // An unmatched field is an author's typo and is reported as one. It is
      // not a blank cell, and treating it as one dropped every row — a merge
      // deleted by a misspelling, under a caution that in the same breath said
      // the placeholder would stay on the slides.
      const misspelled = { ...skipping, slideFields: [["Name"], ["Regoin"], []], fields: ["Name", "Regoin"] };
      expect(skippedRows(misspelled)).toEqual([]);
      const said = caution(misspelled, "merge") ?? "";
      expect(said).toContain("No column for Regoin");
      expect(said, "and does not also say every row is going").not.toContain("will be left out");
    });

    it("stops counting the moment the block moves", () => {
      // The count is read off a SNAPSHOT of the block's slides. Typing in the
      // slide-number boxes runs `blockMoved`, which clears the flat field list
      // — and left the per-slide one standing, so the pane went on dropping a
      // row for a blank in a `{{Note}}` on a slide the state no longer named.
      //
      // Asserted here rather than through the pane, because after a block move
      // the user is on the template step, where the caution does not render at
      // all: a wiring test would pass against the bug.
      const moved = blockMoved(skipping);
      expect(moved.fields, "the flat list goes, as it always did").toEqual([]);
      expect(skippedRows({ ...moved, draft: { from: "8", to: "9" } })).toEqual([]);
    });

    it("says so when the answer leaves nothing to merge", () => {
      const everyRowBlank = toRecordSet([
        ["Name", "Renewal", "Note"],
        ["Ada", "yes", ""],
        ["Cy", "yes", ""],
      ]);
      const nothing = { ...skipping, records: everyRowBlank, rows: everyRowBlank.rows.length };
      expect(blockedReason(nothing, "merge")).toContain("leave the whole row out");
      // Not the unticked sentence, which is a different thing the user did.
      expect(blockedReason(nothing, "merge")).not.toContain("unticked");
    });
  });

  it("warns when the attached pictures have nowhere to go", () => {
    /**
     * The picker is offered from the DATA — a column of file names — because
     * the pane's order is choose slides, paste data, then put the fields on,
     * and at paste time the placeholder may not exist yet. The ENGINE places a
     * picture where a field ASKS for one, `{{Photo|image}}`.
     *
     * Both are right. Nothing put them together: a user could attach three
     * files under "3 pictures named in Photo", press merge, and get slides
     * reading `ada.png` as text, with nothing anywhere having said so.
     */
    const withPhotos = {
      ...state,
      fields: ["Name"],
      images: new Map([["ada.png", new Uint8Array([1])]]),
      imageFields: [],
    };
    // `?? ""` so a caution that goes missing fails as "expected '' to contain
    // …" rather than as an argument-type complaint about null, which says
    // nothing about what broke.
    expect(caution(withPhotos, "merge") ?? "").toContain("will not be placed");
    expect(caution(withPhotos, "merge") ?? "").toContain("{{Column|image}}");

    // A field that asks for a picture: no warning.
    expect(caution({ ...withPhotos, imageFields: ["Photo"] }, "merge") ?? "").not.toContain("will not be placed");
    // Before the slides have been read, an empty list means "not looked yet".
    expect(caution({ ...withPhotos, fields: [] }, "merge") ?? "").not.toContain("will not be placed");
    // And no pictures attached is not a warning about pictures.
    expect(caution({ ...withPhotos, images: undefined }, "merge") ?? "").not.toContain("will not be placed");
  });

  it("says above the button when two picture names differ only by their folder", () => {
    // The last chance to see it: the tally says both matched, because both do
    // resolve to a file that was attached — the same one.
    const clashing: PaneState = {
      ...EMPTY,
      block: { from: 4, to: 6 },
      fields: ["Name", "Photo"],
      imageFields: ["Photo"],
      ...attached("Name,Photo\nAda,eu/logo.png\nBo,us/logo.png"),
      images: new Map([["logo.png", new Uint8Array([1])]]),
    };
    expect(caution(clashing, "merge") ?? "").toContain("same file name in different folders");
    // And it stacks with the other two rather than replacing either.
    const alsoUnmatched = { ...clashing, fields: ["Nmae", "Photo"] };
    const said = caution(alsoUnmatched, "merge") ?? "";
    expect(said).toContain("No column for Nmae");
    expect(said).toContain("same file name in different folders");
  });

  it("is the number on the BUTTON as well as in the sentence", () => {
    /**
     * The half nobody checked. `plannedSlides` was written, covered, and asked
     * by the forecast card — and the button's own label went on multiplying
     * slides-per-record by rows. With this block and these rows the card read
     * one number and the button beside it read another, and the wrong one was
     * on the thing being pressed.
     *
     * The assertion above already says the two answers differ, so it could
     * never have caught this: it pins that `plannedSlides` is not the product
     * without asking whether anything still shows the product.
     */
    const label = primary(state, "merge").label;
    expect(label).toBe(`Add ${plannedSlides(state)} slides`);
    expect(label).not.toBe(`Add ${slidesPerRecord(state.block as Block) * includedCount(state)} slides`);
  });

  it("still counts from the row COUNT when the rows themselves are not in hand", () => {
    // A state can carry `rows` without `records` — the pane knows how many rows
    // it has before it needs them, and `includedCount` is written for exactly
    // that. A condition cannot be evaluated without a row, so the product is
    // the only answer there is, and answering zero would empty the button.
    const counted = { ...state, records: undefined };
    expect(slidesToAdd(counted)).toBe(slidesPerRecord(state.block as Block) * includedCount(counted));
    expect(slidesToAdd(counted)).toBeGreaterThan(0);
  });

  it("counts every slide when nothing is conditional", () => {
    // The other half: the fix must not start subtracting slides nobody skipped.
    const plain = { ...state, conditions: undefined };
    expect(plannedSlides(plain)).toBe(slidesPerRecord(state.block as Block) * includedCount(plain));
  });

  it("is zero before a block is chosen", () => {
    expect(plannedSlides({ fields: [], previewing: false })).toBe(0);
  });
});

describe("what the slide-number boxes accept", () => {
  /**
   * `Number` reads far more than anybody types into a slide box: `0x10` is
   * sixteen, `0b11` is three, `0o17` is fifteen, `1e2` is a hundred.
   *
   * Both halves of that were wrong. `0b11` was ACCEPTED, so the pane quietly
   * offered to merge slides 3 to 9 for somebody who had typed neither number
   * into either box. And the refusals named a cause they had invented — `0x10`
   * produced "The block ends before it starts: slide 16 to 9", about a slide 16
   * that appears nowhere on the user's screen.
   *
   * Third time `Number()` has been wider than what a person typed in this
   * codebase, after `0x10` in a data cell and an unreadable grouping.
   */
  it.each([["0x10"], ["0b11"], ["0o17"], ["1e2"], ["1_0"], ["Infinity"], ["abc"]])(
    "refuses %s, and says so about what was typed",
    (typed) => {
      const read = readBlockDraft({ from: typed, to: "9" });
      expect(read.block, "read as a slide number nobody typed").toBeNull();
      expect(read.why).toBe(`Slide numbers are whole numbers, and "${typed}" is not one.`);
    },
  );

  it("still takes the ordinary spellings of a whole number", () => {
    // The other half. `4.0` IS a whole number, and refusing it with "4.0 is not
    // one" would be a false sentence, so it stays admitted.
    for (const typed of ["4", "4.0", "+4", " 4 "]) {
      expect(readBlockDraft({ from: typed, to: "9" }).block, typed).toEqual({ from: 4, to: 9 });
    }
  });

  it("keeps the refusals that were already specific", () => {
    expect(readBlockDraft({ from: "4.5", to: "9" }).why).toContain('"4.5" is not one');
    expect(readBlockDraft({ from: "-1", to: "9" }).why).toContain("numbered from 1");
    expect(readBlockDraft({ from: "9", to: "4" }).why).toContain("ends before it starts");
    expect(readBlockDraft({ from: "", to: "9" }).why, "red on the first keystroke").toBeNull();
  });
});

describe("what a screen reader is told", () => {
  /**
   * The pane is silent for up to two and a half minutes on an ordinary merge,
   * and until 2026-08-30 it was silent to assistive technology permanently:
   * one `aria` attribute in the whole renderer and no live region at all.
   *
   * Two sentences, and the restraint is the design. The pane redraws on every
   * KEYSTROKE, so anything that moves as the user types interrupts them on
   * every character.
   */
  const idle: PaneState = { fields: [], previewing: false };

  it("says what is out while a host call is out", () => {
    expect(announcement({ ...idle, running: "merge", inFlight: "inserting the merged deck" })).toBe(
      "Waiting on PowerPoint: inserting the merged deck",
    );
  });

  it("says what happened once it is over", () => {
    expect(announcement({ ...idle, notice: "720 slides added after slide 12." })).toBe(
      "720 slides added after slide 12.",
    );
  });

  it("prefers the call in flight to the last run's outcome", () => {
    // Both are set for the whole of a second run: `notice` is what the LAST run
    // said and is still on screen. Announcing it while a new call is out would
    // report a finished run as the current state.
    expect(
      announcement({ ...idle, running: "merge", inFlight: "exporting the template slides", notice: "3 slides added." }),
    ).toBe("Waiting on PowerPoint: exporting the template slides");
  });

  it("says nothing when there is nothing to say", () => {
    // The empty string rather than a word: a live region written with any text
    // announces it, so "idle" has to be silence and not a sentence.
    expect(announcement(idle)).toBe("");
    // Running, but the host has not named a call yet — the window between the
    // press and the first traced call.
    expect(announcement({ ...idle, running: "merge" })).toBe("");
  });

  it("does not announce the headline, which moves on every keystroke", () => {
    // The guard for the restraint above. A draft being typed changes the
    // headline and must not change what is announced.
    const typing: PaneState = { ...idle, draft: { from: "4", to: "" } };
    expect(announcement(typing)).toBe(announcement({ ...typing, draft: { from: "4", to: "6" } }));
  });
});
