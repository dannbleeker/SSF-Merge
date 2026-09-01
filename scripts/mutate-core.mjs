#!/usr/bin/env node
/**
 * Which of this suite's assertions actually discriminate?
 *
 *   node scripts/mutate-core.mjs [max-mutants]
 *
 * `test-kit/driver/mutate.mjs` asks this of the round's VERIFIER — it breaks a
 * merged deck and checks each guard fires. Nothing asked it of the suite, and
 * that gap hid a live defect for the life of a file: `cloneSlide` drew its
 * creation id at random with nothing checking it against the deck, and the test
 * named `gives every copy its own creation id` injected a counter and then
 * asserted the counter's own uniqueness. The generator a real run uses was
 * never in the assertion. See #197.
 *
 * So: change the source in one small way, run the suite, and see whether
 * anything goes red. A mutant that SURVIVES names a behaviour no test is
 * holding — which is a fact about the tests, not necessarily a bug.
 *
 * **Read a survivor before believing it.** Many are equivalent mutants: `x ?? ""`
 * and `x || ""` cannot differ when x is a string, and no test can tell them
 * apart because there is nothing to tell apart. The ones worth acting on are
 * where the two spellings really do behave differently on input a user can
 * produce — a cell holding only spaces, an edit landing exactly on a run
 * boundary. Its first run found both of those, and they are tests now.
 *
 * Not wired into CI, and deliberately: it runs the whole suite once per mutant,
 * which is minutes rather than seconds, and a check that slow in front of every
 * merge gets switched off after the first bad week. The sibling project reaches
 * for a weekly quality sweep for the same reason. Run it when a file's coverage
 * looks good and you do not believe it.
 */
import { execFileSync } from "node:child_process";
import { cpSync, mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { isMain } from "./is-main.mjs";

/** Where a wrong answer costs the most: the engine, and the undo that deletes slides. */
const TARGETS = [
  "src/core/data/format.ts",
  "src/core/pptx/pkg.ts",
  "src/host/undo.ts",
  "src/core/pptx/clone.ts",
  "src/core/merge/numbers.ts",
  "src/core/merge/plan.ts",
  "src/core/merge/resolve.ts",
  "src/core/data/recordset.ts",
  "src/core/merge/text.ts",
  "src/core/merge/images.ts",
];

/**
 * Mechanical edits, each a plausible slip rather than a random byte flip.
 *
 * A boundary moved by one, a guard loosened, a normalisation dropped: the
 * mistakes that survive review because the code still reads correctly.
 */
const RULES = [
  { name: "<= to <", find: /([^<>=!])<= /g, to: (_m, p) => `${p}< ` },
  { name: ">= to >", find: /([^<>=!])>= /g, to: (_m, p) => `${p}> ` },
  { name: "=== to !==", find: / === /g, to: () => " !== " },
  { name: "&& to ||", find: / && /g, to: () => " || " },
  { name: "trim dropped", find: /\.trim\(\)/g, to: () => "" },
  { name: "?? to ||", find: / \?\? /g, to: () => " || " },
];

/**
 * A COPY of the tree to mutate, so the repository is never edited.
 *
 * This wrote its mutants into the working repository and restored them in a
 * `finally`. That holds for a run that fails; it does not hold for one that is
 * KILLED — a `timeout`, a cancelled session, the machine going away — and a
 * skipped restore leaves a mutated engine on disk, which is the outcome the
 * comment on that `finally` calls worse than any finding this script could
 * produce. It happened on 2026-09-01: a run under `timeout` was interrupted, an
 * unrelated `git add -A` swept the live mutant into a commit, and a dropped
 * `.trim()` reached the branch inside a commit about a changelog entry.
 *
 * Copied rather than checked out, so uncommitted work is measured too — which
 * is the state somebody running this is usually asking about. `node_modules` is
 * linked rather than copied: it is most of the bytes, and a junction works on
 * Windows without privileges as well as on POSIX.
 *
 * `.git` IS copied, and that is not incidental. It was left out at first — it
 * is 27M and no mutant in `src/core` could reach a test about committed files —
 * and the omission made the whole tool vacuous: `test/docprops.test.ts` calls
 * `git ls-files` at collection time, which throws outside a repository, so the
 * copy's suite was red before any mutant was written. Every mutant then read
 * "caught", `survivors` stayed empty, and the run ended on a line that looked
 * like a perfect score. A commit shipped claiming that control had been checked;
 * it had not, and the check that would have caught it read `tail`'s exit code
 * rather than vitest's.
 *
 * So the control is RUN, here, by this script, and it is not advice to the
 * reader: `main` copies the tree, runs the suite with nothing mutated, and
 * refuses to report survivors at all if that comes back red. A sweep whose
 * baseline is red cannot distinguish a killed mutant from a broken copy, and
 * the failure mode is a clean-looking result rather than an error.
 */
function workingCopy() {
  const dir = mkdtempSync(join(tmpdir(), "ssf-mutate-"));
  cpSync(".", dir, {
    recursive: true,
    filter: (from) => !/(^|[\\/])(node_modules|dist|dist-lib|coverage)([\\/]|$)/.test(from),
  });
  symlinkSync(join(process.cwd(), "node_modules"), join(dir, "node_modules"), "junction");
  return dir;
}

/**
 * The suite, run in the copy.
 *
 * Returns the output as well as the verdict, because the CONTROL has to be able
 * to say why it is red. A bare boolean is what let a broken copy read as a
 * flawless sweep.
 *
 * @param {string} cwd
 * @returns {{ green: boolean; output: string }}
 */
function runSuite(cwd) {
  try {
    const out = execFileSync("node", ["./node_modules/vitest/vitest.mjs", "run", "--silent"], {
      cwd,
      stdio: "pipe",
      encoding: "utf8",
      timeout: 600000,
    });
    return { green: true, output: out };
  } catch (e) {
    const err = /** @type {{ stdout?: string; stderr?: string; message?: string }} */ (e);
    return { green: false, output: `${err.stdout ?? ""}${err.stderr ?? ""}` || (err.message ?? "") };
  }
}

/**
 * Whether this hit is inside a comment.
 *
 * A mutant in a doc comment changes nothing, so the suite stays green and the
 * run reports a SURVIVOR that names no gap at all — one run's list held
 * `plan.ts:242 ?? to || — * Only a field the data has a COLUMN for`, which is
 * prose. The middle-hit rule was chosen precisely to avoid a doc comment's
 * example and it only moved the problem along.
 *
 * @param {string} source @param {number} at
 */
function inComment(source, at) {
  const line = (source.slice(0, at).split("\n").pop() ?? "") + (source.slice(at).split("\n")[0] ?? "");
  return /^\s*(\*|\/\/|\/\*)/.test(line);
}

/**
 * How many tests the control ran, read out of vitest's own summary.
 *
 * Printed so a reader can compare it with a plain `vitest run`. A copy that
 * quietly collects fewer files is the same vacuous measurement as a red one,
 * one step less obvious.
 *
 * @param {string} output
 */
function countedTests(output) {
  const m = /Tests\s+(\d+) passed/.exec(output);
  return m ? Number(m[1]) : undefined;
}

function main() {
  const cap = Number(process.argv[2] ?? 40);
  /** @type {string[]} */
  const survivors = [];
  let applied = 0;
  const copy = workingCopy();
  console.log(`mutating a copy at ${copy}`);

  try {
    // The control, run rather than recommended. Nothing is mutated yet, so this
    // is the copy answering for itself: if it is red here, every mutant below
    // reads "caught" and the run reports a flawless sweep it has not measured.
    const control = runSuite(copy);
    if (!control.green) {
      console.log("the copy's suite is RED with nothing mutated — no sweep is possible from here");
      console.log(control.output.split("\n").slice(-25).join("\n"));
      process.exitCode = 1;
      return;
    }
    // The count as well as the colour. A copy that collects FEWER test files
    // and is green is the same vacuous sweep one step less obvious, and the
    // repo already keeps the number to check it against — `test-count.mjs`
    // holds a floor that rises on its own. Printed and asserted, not printed
    // and left for a reader to notice.
    const ran = countedTests(control.output);
    const floor = JSON.parse(readFileSync("test/fixtures/test-count.json", "utf8")).min;
    if (ran === undefined || ran < floor) {
      console.log(`the copy ran ${ran ?? "an unreadable number of"} test(s) against a floor of ${floor}`);
      console.log("a copy that collects less than the repo does cannot answer for the repo");
      process.exitCode = 1;
      return;
    }
    console.log(`control: the unmutated copy is green — ${ran} test(s), floor ${floor}`);

    outer: for (const file of TARGETS) {
      const original = readFileSync(file, "utf8");
      for (const rule of RULES) {
        // Code only. A mutant in a doc comment changes nothing, so it survives
        // and names a gap that is not there.
        const hits = [...original.matchAll(rule.find)].filter((h) => !inComment(original, h.index ?? 0));
        if (!hits.length) continue;
        // One site per rule per file: enough to ask the question without a
        // combinatorial sweep nobody would wait for. The middle hit rather than
        // the first, because the first is often in a doc comment's example.
        const hit = hits[Math.floor(hits.length / 2)];
        const at = hit.index ?? 0;
        const mutated = original.slice(0, at) + rule.to(hit[0], hit[1] ?? "") + original.slice(at + hit[0].length);
        if (mutated === original) continue;

        writeFileSync(join(copy, file), mutated);
        let green;
        try {
          green = runSuite(copy).green;
        } finally {
          // Restored inside the COPY. Nothing here can leave the repository
          // mutated, however this run ends.
          writeFileSync(join(copy, file), original);
        }
        applied++;
        const line = original.slice(0, at).split("\n").length;
        console.log(`${green ? "SURVIVED " : "caught   "} ${file}:${line}  ${rule.name}`);
        if (green)
          survivors.push(
            `${file}:${line} ${rule.name} — ${(original.split("\n")[line - 1] ?? "").trim().slice(0, 90)}`,
          );
        if (applied >= cap) break outer;
      }
    }
  } finally {
    rmSync(copy, { recursive: true, force: true });
  }

  console.log(`\n${applied} mutant(s), ${survivors.length} survived. Read each before believing it.`);
  for (const s of survivors) console.log("  " + s);
}

if (isMain(import.meta.url)) main();
