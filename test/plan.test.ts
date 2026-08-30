import { describe, expect, it } from "vitest";
import { toRecordSet } from "../src/core/data/recordset.js";
import { buildPlan, isTruthy, recordCount, slideCount, type Block } from "../src/core/merge/plan.js";
import { TAG_BLOCK, TAG_RECORD, TAG_RUN, TAG_SEQ } from "../src/core/pptx/tags.js";

const records = toRecordSet([
  ["Name", "Notes"],
  ["Ada", "a note"],
  ["Grace", ""],
  ["Katherine", "another"],
]);

/** Three slides, the third only when Notes is filled. */
const block: Block = {
  id: "block-a",
  slides: [
    { path: "ppt/slides/slide4.xml", seq: 1 },
    { path: "ppt/slides/slide5.xml", seq: 2 },
    { path: "ppt/slides/slide6.xml", seq: 3, condition: "Notes" },
  ],
};

const opts = { runId: "run-fixed" };

describe("buildPlan", () => {
  it("emits every slide of one record before starting the next", () => {
    // Record-major. A record is the thing the deck is about, so its slides stay
    // together; all the covers followed by all the detail pages is a different
    // deck nobody asked for.
    const plan = buildPlan(block, records, opts);
    expect(plan.steps.map((s) => `${s.recordIndex}:${s.seq}`)).toEqual([
      "0:1",
      "0:2",
      "0:3",
      "1:1",
      "1:2",
      "2:1",
      "2:2",
      "2:3",
    ]);
  });

  it("skips a conditional slide in place, leaving the others' order alone", () => {
    const plan = buildPlan(block, records, opts);
    const grace = plan.steps.filter((s) => s.recordIndex === 1);
    expect(grace.map((s) => s.seq)).toEqual([1, 2]);
    expect(plan.skippedSlides).toEqual([{ recordIndex: 1, seq: 2 + 1, condition: "Notes" }]);
  });

  it("tags every step with the run, the block, the sequence and the record", () => {
    // These go into the file before it reaches PowerPoint. They are what makes
    // undo and re-run possible on a host that will not resolve a fresh slide.
    const plan = buildPlan(block, records, opts);
    expect(plan.steps[0]?.tags).toEqual([
      [TAG_RUN, "run-fixed"],
      [TAG_BLOCK, "block-a"],
      [TAG_SEQ, "1"],
      [TAG_RECORD, "0"],
    ]);
  });

  it("reads the block in the deck's order, not the order it was handed", () => {
    const shuffled: Block = { id: "b", slides: [block.slides[2]!, block.slides[0]!, block.slides[1]!] };
    const plan = buildPlan(shuffled, records, opts);
    expect(plan.steps.filter((s) => s.recordIndex === 0).map((s) => s.seq)).toEqual([1, 2, 3]);
  });

  it("merges only the rows it was given", () => {
    const plan = buildPlan(block, records, { ...opts, recordIndexes: [2] });
    expect(recordCount(plan)).toBe(1);
    expect(plan.steps.every((s) => s.recordIndex === 2)).toBe(true);
  });

  it("drops a whole record when a bound field is empty and the policy says skip", () => {
    const withFields: Block = {
      id: "b",
      slides: [{ path: "ppt/slides/slide4.xml", seq: 1, fields: ["Notes"] }],
    };
    const plan = buildPlan(withFields, records, { ...opts, onEmpty: "skip" });
    expect(plan.skippedRecords).toEqual([1]);
    expect(slideCount(plan)).toBe(2);
  });

  it("emits the slide when a condition names a column that does not exist, and reports it", () => {
    // Dropping it would hide an authoring mistake behind output that looks
    // finished. The slide is produced and the pane is told why.
    const bad: Block = { id: "b", slides: [{ path: "ppt/slides/slide4.xml", seq: 1, condition: "Territory" }] };
    const plan = buildPlan(bad, records, opts);
    expect(slideCount(plan)).toBe(3);
    expect(plan.unknownConditions).toEqual(["Territory"]);
  });

  it("answers an empty plan for an empty record set rather than throwing", () => {
    const plan = buildPlan(block, toRecordSet([["Name"]]), opts);
    expect(plan.steps).toEqual([]);
  });

  it("refuses a block with no slides, which is a programming error", () => {
    expect(() => buildPlan({ id: "b", slides: [] }, records, opts)).toThrow(/at least one slide/);
  });
});

describe("isTruthy", () => {
  it("reads a blank cell as false", () => {
    expect(isTruthy("")).toBe(false);
    expect(isTruthy("   ")).toBe(false);
    expect(isTruthy(undefined)).toBe(false);
  });

  it("reads a spreadsheet's written-out negatives as false, in English and Danish", () => {
    // Excel exports a boolean as a localised WORD, so a merge that treated
    // "FALSK" as content would emit every slide it was told to leave out.
    for (const word of ["false", "FALSE", "No", "nej", "FALSK", "off", "0"]) {
      expect(isTruthy(word), word).toBe(false);
    }
  });

  it("reads anything else as true, including a word that merely starts with n", () => {
    for (const word of ["yes", "sand", "1", "note", "Nordics", "0.0"]) {
      expect(isTruthy(word), word).toBe(true);
    }
  });
});

describe("a condition decides before an empty cell does", () => {
  /**
   * `onEmpty: "skip"` drops a record whose fields are not all filled. It read
   * the fields of EVERY slide in the block, including ones the record's own
   * conditions had already left out — so a customer with no renewal note
   * vanished from the deck over a blank cell on the renewal slide they were
   * never going to get.
   *
   * No longer latent: the pane sets the policy now, through the blank-cell
   * control on the merge step, and `plannedSlides` asks the same rule so the
   * number on the button follows. The failure this guards is a record silently
   * absent.
   */
  const block: Block = {
    id: "b",
    slides: [
      { path: "s1.xml", seq: 1, fields: ["Name"] },
      { path: "s2.xml", seq: 2, condition: "HasExtra", fields: ["ExtraNote"] },
    ],
  };
  const records = toRecordSet([
    ["Name", "HasExtra", "ExtraNote"],
    ["Ada", "no", ""],
    ["Bo", "yes", "a note"],
  ]);

  it("keeps a record whose blank field is on a slide it is not getting", () => {
    const plan = buildPlan(block, records, { runId: "r", onEmpty: "skip" });
    expect(plan.skippedRecords, "the record was dropped over a slide it never gets").toEqual([]);
    expect(plan.steps.map((s) => `${s.recordIndex}/${s.seq}`)).toEqual(["0/1", "1/1", "1/2"]);
  });

  it("still reports that slide as left out by its condition", () => {
    // The record is present and one of its slides is not, which is two facts
    // and both have to be said.
    const plan = buildPlan(block, records, { runId: "r", onEmpty: "skip" });
    expect(plan.skippedSlides).toEqual([{ recordIndex: 0, seq: 2, condition: "HasExtra" }]);
  });

  it("still drops a record whose blank field is on a slide it IS getting", () => {
    // The policy itself is unchanged, which is the other half: narrowing what
    // it looks at must not stop it looking.
    const blanks = toRecordSet([
      ["Name", "HasExtra", "ExtraNote"],
      ["", "yes", "a note"],
      ["Bo", "yes", ""],
    ]);
    const plan = buildPlan(block, blanks, { runId: "r", onEmpty: "skip" });
    expect(plan.skippedRecords).toEqual([0, 1]);
    expect(plan.steps).toEqual([]);
  });

  it("says nothing about the slides of a record it dropped whole", () => {
    // A record that contributed nothing reporting two absences as well would
    // be two answers about one record.
    const dropped = toRecordSet([
      ["Name", "HasExtra", "ExtraNote"],
      ["", "no", ""],
    ]);
    const plan = buildPlan(block, dropped, { runId: "r", onEmpty: "skip" });
    expect(plan.skippedRecords).toEqual([0]);
    expect(plan.skippedSlides).toEqual([]);
  });
});

describe("only a field that is actually on a slide, and actually a column", () => {
  /**
   * Two ways `onEmpty: "skip"` could drop a record over something that is not
   * a blank cell in a field the merge uses. Both matter to the number on the
   * merge button, and both would be silent: a record simply absent.
   *
   * A spreadsheet routinely carries columns a template does not use, and half
   * of them are empty — an unfilled "Comment", a column somebody stopped
   * maintaining. Dropping a row for one of those would make the control
   * useless on real data.
   */
  const block: Block = {
    id: "b",
    slides: [
      { path: "s1.xml", seq: 1, fields: ["Name"] },
      { path: "s2.xml", seq: 2, fields: ["Notes"] },
    ],
  };
  const records = toRecordSet([
    ["Name", "Notes", "Comment"],
    ["Ada", "kept", ""],
    ["Bo", "kept", ""],
    ["Cy", "kept", "said something"],
  ]);

  it("ignores a blank column no slide refers to", () => {
    const plan = buildPlan(block, records, { runId: "r", onEmpty: "skip" });
    expect(plan.skippedRecords).toEqual([]);
    expect(plan.steps).toHaveLength(6);
  });

  it("drops those same rows once that column IS a field, so this is not vacuous", () => {
    const usingComment: Block = {
      id: "b",
      slides: [
        { path: "s1.xml", seq: 1, fields: ["Name"] },
        { path: "s2.xml", seq: 2, fields: ["Comment"] },
      ],
    };
    expect(buildPlan(usingComment, records, { runId: "r", onEmpty: "skip" }).skippedRecords).toEqual([0, 1]);
  });

  it("ignores a field the data has no column for at all", () => {
    // `row[field] ?? ""` cannot tell a blank cell from a column that is not
    // there, and one misspelled placeholder would otherwise drop EVERY row —
    // the merge deleted by a typo, under a caution saying in the same breath
    // that the placeholder will stay on the slides, about slides no row would
    // produce. An unmatched field is reported as one; it is not data.
    const misspelled: Block = {
      id: "b",
      slides: [
        { path: "s1.xml", seq: 1, fields: ["Name"] },
        { path: "s2.xml", seq: 2, fields: ["Regoin"] },
      ],
    };
    expect(buildPlan(misspelled, records, { runId: "r", onEmpty: "skip" }).skippedRecords).toEqual([]);
  });
});
