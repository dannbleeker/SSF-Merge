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
import { readFileSync, writeFileSync } from "node:fs";
import { isMain } from "./is-main.mjs";

/** Where a wrong answer costs the most: the engine, and the undo that deletes slides. */
const TARGETS = [
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

function suiteGreen() {
  try {
    execFileSync("node", ["./node_modules/vitest/vitest.mjs", "run", "--silent"], { stdio: "pipe", timeout: 600000 });
    return true;
  } catch {
    return false;
  }
}

function main() {
  const cap = Number(process.argv[2] ?? 40);
  /** @type {string[]} */
  const survivors = [];
  let applied = 0;

  outer: for (const file of TARGETS) {
    const original = readFileSync(file, "utf8");
    for (const rule of RULES) {
      const hits = [...original.matchAll(rule.find)];
      if (!hits.length) continue;
      // One site per rule per file: enough to ask the question without a
      // combinatorial sweep nobody would wait for. The middle hit rather than
      // the first, because the first is often in a doc comment's example.
      const hit = hits[Math.floor(hits.length / 2)];
      const at = hit.index ?? 0;
      const mutated = original.slice(0, at) + rule.to(hit[0], hit[1] ?? "") + original.slice(at + hit[0].length);
      if (mutated === original) continue;

      writeFileSync(file, mutated);
      let green;
      try {
        green = suiteGreen();
      } finally {
        // Restored even if the run throws: leaving a mutated engine on disk is
        // a worse outcome than any finding this script could produce.
        writeFileSync(file, original);
      }
      applied++;
      const line = original.slice(0, at).split("\n").length;
      console.log(`${green ? "SURVIVED " : "caught   "} ${file}:${line}  ${rule.name}`);
      if (green)
        survivors.push(`${file}:${line} ${rule.name} — ${(original.split("\n")[line - 1] ?? "").trim().slice(0, 90)}`);
      if (applied >= cap) break outer;
    }
  }

  console.log(`\n${applied} mutant(s), ${survivors.length} survived. Read each before believing it.`);
  for (const s of survivors) console.log("  " + s);
}

if (isMain(import.meta.url)) main();
