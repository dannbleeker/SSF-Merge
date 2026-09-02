import { describe, expect, it } from "vitest";
// @ts-expect-error — a plain .mjs tool with no types, shared with the scripts.
import { skippedNames as untypedNames, verdict as untypedVerdict } from "../scripts/test-count.mjs";

interface Record_ {
  min: number;
  maxSkipped?: number;
}
interface Answer {
  ok: boolean;
  message: string;
  write?: { min: number; maxSkipped: number } | null;
}
const verdict = untypedVerdict as (o: {
  defined: number;
  skipped: number;
  record: Record_;
  update?: boolean;
}) => Answer;
const skippedNames = untypedNames as (report: unknown) => string[];

/**
 * The gate that holds a floor under the suite, which had no test of its own.
 *
 * It is the one gate whose failure is invisible: every other check fails loudly
 * when it breaks, and this one just stops noticing. It has been wrong twice —
 * once counting pending tests, so 23 could be switched off in silence, and once
 * counting only tests that RAN, which made the number depend on the machine.
 */
describe("the floor under the suite", () => {
  const record = { min: 100, maxSkipped: 1 };

  it("passes a run that matches", () => {
    expect(verdict({ defined: 100, skipped: 0, record }).ok).toBe(true);
  });

  it("raises the floor on its own when the suite grows", () => {
    const answer = verdict({ defined: 120, skipped: 0, record });
    expect(answer.ok).toBe(true);
    expect(answer.write).toEqual({ min: 120, maxSkipped: 1 });
  });

  it("refuses a suite that has lost tests", () => {
    const answer = verdict({ defined: 77, skipped: 0, record });
    expect(answer.ok).toBe(false);
    expect(answer.message).toContain("down from 100");
  });

  it("refuses 23 tests switched off, which is what it is for", () => {
    // `it(` to `it.skip(` leaves the total alone, so the floor cannot see it.
    // The cap is what does, and this is the case that motivated the whole gate:
    // the merge plan's decision engine, silent, every check green.
    const answer = verdict({ defined: 100, skipped: 23, record });
    expect(answer.ok).toBe(false);
    expect(answer.message).toContain("23 tests are skipped");
  });

  it("is the same number on a machine that must skip one", () => {
    /**
     * The reason it counts tests that EXIST rather than tests that ran.
     *
     * `is-main.test.ts` needs a symlink and Windows refuses that without
     * elevation. Counting only what ran gave 1475 there and 1476 in CI, and
     * either number committed breaks the other machine.
     */
    const ci = verdict({ defined: 1476, skipped: 0, record: { min: 1476, maxSkipped: 1 } });
    const windows = verdict({ defined: 1476, skipped: 1, record: { min: 1476, maxSkipped: 1 } });
    expect(ci.ok, "CI runs every test").toBe(true);
    expect(windows.ok, "and the machine that cannot make a symlink agrees").toBe(true);
    expect(windows.write, "neither rewrites the record").toBeNull();
    expect(ci.write).toBeNull();
  });

  it("does not raise the skip cap by itself", () => {
    // A suite growing is ordinary; a new skip is a decision somebody should be
    // seen making. Only `--update` moves this one.
    const answer = verdict({ defined: 130, skipped: 0, record });
    expect(answer.write).toEqual({ min: 130, maxSkipped: 1 });
  });

  it("records a deliberate drop, and a deliberate skip, only with --update", () => {
    const answer = verdict({ defined: 90, skipped: 4, record, update: true });
    expect(answer.ok).toBe(true);
    expect(answer.write).toEqual({ min: 90, maxSkipped: 4 });
  });

  it("treats a record with no cap as allowing none", () => {
    // Written before the cap existed. Infinity would silently accept whatever
    // was skipped on the day somebody upgraded.
    const answer = verdict({ defined: 100, skipped: 1, record: { min: 100 } });
    expect(answer.ok).toBe(false);
    expect(answer.message).toContain("allows 0");
  });

  it("refuses a count that is not a count rather than comparing it", () => {
    for (const defined of [NaN, 0, -1, 1.5, undefined as unknown as number]) {
      expect(verdict({ defined, skipped: 0, record }).ok, `defined ${defined}`).toBe(false);
    }
    for (const skipped of [NaN, -1, 1.5, undefined as unknown as number]) {
      expect(verdict({ defined: 100, skipped, record }).ok, `skipped ${skipped}`).toBe(false);
    }
  });
});

describe("naming what was skipped", () => {
  it("lists them, so the log says which rather than how many", () => {
    const report = {
      testResults: [
        {
          assertionResults: [
            { status: "passed", fullName: "a passing one" },
            { status: "pending", fullName: "whether a module is the entry point > through a SYMLINK" },
          ],
        },
        { assertionResults: [{ status: "skipped", fullName: "another" }] },
      ],
    };
    expect(skippedNames(report)).toEqual(["whether a module is the entry point > through a SYMLINK", "another"]);
  });

  it("answers nothing for a report with no results", () => {
    expect(skippedNames({})).toEqual([]);
    expect(skippedNames(undefined)).toEqual([]);
  });
});
