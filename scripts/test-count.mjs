#!/usr/bin/env node
/**
 * Hold a floor under the number of tests.
 *
 * Moving tests between files is where suites lose cases silently: the run stays
 * green because the deleted tests are simply not there to fail. A sibling
 * project lost 43 that way and found out much later. Comparing the total
 * against a recorded number costs one file and catches it in the diff.
 *
 * The floor rises on its own when the suite grows, because a floor that only
 * moves by hand drifts behind and stops catching a partial deletion. A
 * deliberate DROP is re-recorded with `--update`, so it lands in a commit where
 * a reviewer sees it rather than being absorbed silently.
 *
 * **Two numbers, not one, and the second is why.** This used to count only
 * tests that RAN — `numTotalTests` minus the pending ones — because counting
 * pending tests is what let 23 of them be switched off (`it(` to `it.skip(`)
 * with every gate green: the whole merge-plan decision engine, silent, floor
 * untouched.
 *
 * That was right about the danger and wrong about the measure. A test skipped
 * because the MACHINE cannot run it is not a test somebody switched off, and
 * counting the two the same way made the floor platform-dependent: this repo's
 * `is-main.test.ts` needs a symlink, Windows refuses that without elevation, so
 * the same commit counted 1475 on Windows and 1476 on Linux. Either number
 * committed breaks the other machine — and the first way out taken was worse
 * than the problem, because it left the whole suite failing locally, which
 * teaches everyone to scroll past a red run.
 *
 * So the floor now counts tests that EXIST, which is the same number on every
 * machine, and a second recorded number caps how many may be skipped. Switching
 * off 23 tests leaves the total alone and blows the cap; deleting 23 drops the
 * total. Both still caught, and neither number depends on who ran it.
 */
import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { isMain } from "./is-main.mjs";

const RECORD = "test/fixtures/test-count.json";

/**
 * What to do about a run, decided from the numbers alone.
 *
 * Separated from running vitest so it can be tested without running the suite
 * inside itself. Every branch below has cost this repository something, and
 * none of them was provable while this was a column of `if`s wrapped around an
 * `execFileSync`.
 */
export function verdict({ defined, skipped, record, update = false }) {
  const min = Number.isInteger(record?.min) ? record.min : 0;
  // Absent in a record written before this change. Zero rather than Infinity: a
  // repository that has never recorded a cap has never agreed to a skip.
  const maxSkipped = Number.isInteger(record?.maxSkipped) ? record.maxSkipped : 0;

  if (!Number.isInteger(defined) || defined <= 0) {
    return { ok: false, message: `the run reported ${defined} tests, which is not a count. Refusing to compare.` };
  }
  if (!Number.isInteger(skipped) || skipped < 0) {
    return { ok: false, message: `the run reported ${skipped} skipped, which is not a count. Refusing to compare.` };
  }

  if (defined < min && !update) {
    return {
      ok: false,
      message:
        `the suite DEFINES ${defined} tests, down from ${min}.\n` +
        `If you deliberately removed tests, re-record it with:  node scripts/test-count.mjs --update`,
    };
  }
  if (skipped > maxSkipped && !update) {
    return {
      ok: false,
      message:
        `${skipped} tests are skipped and the record allows ${maxSkipped}.\n` +
        `A test skipped because this machine cannot run it is fine, and is recorded with:  node scripts/test-count.mjs --update\n` +
        `A test skipped to turn a red run green is the thing this gate exists for.`,
    };
  }

  // The floor rises on its own; the skip cap does NOT. A suite growing is
  // ordinary. A new skip is a decision somebody should be seen making.
  const raise = defined > min;
  return {
    ok: true,
    write:
      update || raise
        ? { min: update ? defined : Math.max(defined, min), maxSkipped: update ? skipped : maxSkipped }
        : null,
    message: raise
      ? `floor raised to ${defined}. Commit ${RECORD}.`
      : update
        ? `re-recorded: ${defined} tests, ${skipped} skipped, deliberately.`
        : `${defined} tests defined, floor ${min}.${skipped > 0 ? ` ${skipped} skipped, cap ${maxSkipped}.` : ""}`,
  };
}

/**
 * The names of the skipped tests, so a reviewer sees WHICH ones in the log.
 *
 * A cap says how many; a reviewer needs to know which, because "1 skipped" is
 * the same line whether it is the symlink test this machine cannot run or a
 * case somebody switched off to get a green run.
 *
 * @param {unknown} report
 * @returns {string[]}
 */
export function skippedNames(report) {
  /** @type {string[]} */
  const out = [];
  const files = /** @type {{ assertionResults?: { status?: string, fullName?: string, title?: string }[] }[]} */ (
    /** @type {{ testResults?: unknown }} */ (report ?? {}).testResults ?? []
  );
  for (const file of files) {
    for (const test of file.assertionResults ?? []) {
      if (test.status === "pending" || test.status === "skipped") out.push(test.fullName ?? test.title ?? "?");
    }
  }
  return out;
}

function main() {
  const update = process.argv.includes("--update");
  const out = join(mkdtempSync(join(tmpdir(), "ssf-merge-")), "results.json");
  // Vitest's own entry point through the running Node, not `npx`. `execFileSync`
  // does not go through a shell, and on Windows the executable is `npx.cmd`, so
  // spawning `npx` there is `ENOENT` before any test runs — this gate could not
  // be run at all on the owner's machine while CI on ubuntu passed it. Naming
  // the installed entry also guarantees the pinned vitest rather than whatever
  // `npx` would resolve, which is what a floor under the suite wants anyway.
  execFileSync(
    process.execPath,
    [join("node_modules", "vitest", "vitest.mjs"), "run", "--reporter=json", `--outputFile=${out}`],
    { stdio: "inherit" },
  );

  const report = JSON.parse(readFileSync(out, "utf8"));
  const record = JSON.parse(readFileSync(RECORD, "utf8"));
  const skipped = Number.isInteger(report.numPendingTests) ? report.numPendingTests : 0;
  const answer = verdict({ defined: report.numTotalTests, skipped, record, update });

  const names = skippedNames(report);
  if (names.length) console.log(`test-count: skipped — ${names.join("; ")}`);

  if (!answer.ok) {
    console.error(`test-count: ${answer.message}`);
    process.exit(1);
  }
  if (answer.write) writeFileSync(RECORD, `${JSON.stringify(answer.write, null, 2)}\n`);
  console.log(`test-count: ${answer.message}`);
}

if (isMain(import.meta.url)) main();
