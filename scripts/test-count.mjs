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
 */
import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const RECORD = "test/fixtures/test-count.json";
const update = process.argv.includes("--update");

const out = join(mkdtempSync(join(tmpdir(), "ssf-merge-")), "results.json");
// Vitest's own entry point through the running Node, not `npx`. `execFileSync`
// does not go through a shell, and on Windows the executable is `npx.cmd`, so
// spawning `npx` there is `ENOENT` before any test runs — this gate could not
// be run at all on the owner's machine while CI on ubuntu passed it. Naming the
// installed entry also guarantees the pinned vitest rather than whatever `npx`
// would resolve, which is what a floor under the suite wants anyway.
execFileSync(
  process.execPath,
  [join("node_modules", "vitest", "vitest.mjs"), "run", "--reporter=json", `--outputFile=${out}`],
  {
    stdio: "inherit",
  },
);

const total = JSON.parse(readFileSync(out, "utf8")).numTotalTests;
const recorded = JSON.parse(readFileSync(RECORD, "utf8")).min;

if (!Number.isInteger(total) || total <= 0) {
  console.error(`test-count: the run reported ${total} tests, which is not a count. Refusing to compare.`);
  process.exit(1);
}

if (total < recorded) {
  if (!update) {
    console.error(
      `test-count: the suite has ${total} tests, down from ${recorded}.\n` +
        `If you deliberately removed tests, re-record it with:  node scripts/test-count.mjs --update`,
    );
    process.exit(1);
  }
  console.log(`test-count: recorded floor lowered from ${recorded} to ${total}, deliberately.`);
}

if (total > recorded || (update && total !== recorded)) {
  writeFileSync(RECORD, `${JSON.stringify({ min: total }, null, 2)}\n`);
  if (total > recorded) console.log(`test-count: floor raised to ${total}. Commit ${RECORD}.`);
} else {
  console.log(`test-count: ${total} tests, floor ${recorded}.`);
}
