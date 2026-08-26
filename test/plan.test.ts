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
