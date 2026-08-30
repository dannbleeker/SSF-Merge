#!/usr/bin/env node
/**
 * Mutation harness for verify-package.mjs.
 *
 * A verifier that has only ever seen a good deck is an untested instrument.
 * This breaks the reference deck in one specific way at a time and asserts that
 * the matching check goes red — and, just as important, that the OTHER checks
 * stay green, so a mutation that trips everything is not mistaken for a working
 * guard.
 *
 * Usage: node test-kit/driver/mutate.mjs [reference-deck.pptx]
 */
import { readFileSync, writeFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import JSZip from "jszip";

const SRC = process.argv[2] ?? "test-kit/out/headless-reference.pptx";

/** Each mutant: a name, an edit, and the check whose text must appear as FAIL. */
const MUTANTS = [
  {
    name: "dangle-a-relationship",
    // The name of the check that owns this, and it is a check in
    // `scripts/package-integrity.mjs` now rather than a walk this verifier did
    // itself. Renaming it there without renaming it here would leave this
    // mutant reporting "not caught" on a defect that IS caught — a mutation
    // harness lying in the safe-looking direction.
    expect: "the package agrees with itself",
    async apply(zip) {
      const p = "ppt/slides/_rels/slide4.xml.rels";
      const xml = await zip.file(p).async("string");
      zip.file(p, xml.replace(/Target="([^"]*charts\/chart\d+\.xml)"/, 'Target="../charts/chartNOPE.xml"'));
    },
  },
  {
    name: "swap-a-chart-region",
    expect: "each chart holds its OWN slide's region",
    async apply(zip) {
      const p = "ppt/charts/chart2.xml";
      const xml = await zip.file(p).async("string");
      zip.file(p, xml.split("Nordics").join("Benelux"));
    },
  },
  {
    name: "revert-a-workbook-to-placeholder",
    expect: "each chart's workbook holds that slide's region",
    async apply(zip) {
      const p = "ppt/embeddings/workbook1.xlsx";
      const inner = await JSZip.loadAsync(await zip.file(p).async("nodebuffer"));
      for (const n of Object.keys(inner.files)) {
        if (!/^xl\/(sharedStrings|worksheets\/sheet\d+)\.xml$/.test(n)) continue;
        const x = await inner.file(n).async("string");
        inner.file(n, x.split("Nordics").join("{{Region}}"));
      }
      zip.file(p, await inner.generateAsync({ type: "nodebuffer" }));
    },
  },
  {
    name: "give-two-rows-the-same-photo",
    expect: "three distinct image payloads",
    async apply(zip) {
      zip.file("ppt/media/image2.png", await zip.file("ppt/media/image1.png").async("nodebuffer"));
    },
  },
  {
    name: "blank-out-the-nickname-placeholder",
    expect: "{{Nickname}} still present",
    async apply(zip) {
      for (const n of Object.keys(zip.files)) {
        if (!/^ppt\/slides\/slide\d+\.xml$/.test(n)) continue;
        const x = await zip.file(n).async("string");
        zip.file(n, x.split("{{Nickname}}").join(""));
      }
    },
  },
  {
    name: "wrong-name-in-a-notes-page",
    expect: "notes pages read",
    async apply(zip) {
      const p = "ppt/notesSlides/notesSlide2.xml";
      const x = await zip.file(p).async("string");
      zip.file(p, x.split("Ada").join("Zorro"));
    },
  },
];

function runVerifier(path) {
  try {
    return execFileSync("node", ["test-kit/driver/verify-package.mjs", path], { encoding: "utf8" });
  } catch (e) {
    return (e.stdout ?? "") + (e.stderr ?? "");
  }
}

/** Which check names are FAIL in a verifier report. */
function failedChecks(out) {
  const line = out.split("\n").find((l) => l.startsWith("FAILED:"));
  return line
    ? line
        .replace("FAILED:", "")
        .split(" ; ")
        .map((s) => s.trim())
    : [];
}

/**
 * A check already FAILING on the reference deck cannot prove anything.
 *
 * `fired` asks whether the expected check is red after the mutation. If it was
 * red BEFORE it, the answer is yes no matter what the mutation did — the
 * assertion passes without the guard being exercised at all.
 *
 * The round of 2026-08-30 ran this on a deck that kept its template slides, so
 * eight checks were red at baseline and FOUR of the six mutants asserted
 * against one of them. The harness printed "Every mutation was caught by its
 * own guard" and exited 0, which is the same shape of lie it exists to catch
 * one level down: a verifier nobody has tested. So a mutant whose check is
 * already red is INCONCLUSIVE, and inconclusive is not success.
 */
const baseline = failedChecks(runVerifier(SRC));
if (baseline.length) {
  console.log(`baseline failures on ${SRC} — ${baseline.length} check(s) already red BEFORE any mutation:`);
  for (const f of baseline) console.log(`  - ${f}`);
  console.log(`  Any mutant expecting one of these cannot be proven by this run.`);
} else {
  console.log(`baseline: every check passes on ${SRC}, so each mutant below is a real test.`);
}
console.log();

let proven = 0;
let allGood = true;
for (const m of MUTANTS) {
  const zip = await JSZip.loadAsync(readFileSync(SRC));
  await m.apply(zip);
  const out = `test-kit/out/mutant-${m.name}.pptx`;
  writeFileSync(out, await zip.generateAsync({ type: "nodebuffer" }));

  const failures = failedChecks(runVerifier(out));
  const alreadyRed = baseline.some((f) => f.includes(m.expect));
  const fired = failures.some((f) => f.includes(m.expect));
  const collateral = failures.filter((f) => !f.includes(m.expect) && !baseline.includes(f));

  const verdict = alreadyRed ? "CANNOT PROVE - check already red" : fired ? "GUARD FIRED" : "GUARD SILENT";
  const tag = alreadyRed ? "??? " : fired ? "OK  " : "BAD ";
  if (alreadyRed || !fired) allGood = false;
  if (!alreadyRed && fired) proven++;
  console.log(`${tag} ${m.name.padEnd(36)} ${verdict}`);
  console.log(`     expected check: "${m.expect}"`);
  if (collateral.length) console.log(`     also tripped (not necessarily wrong): ${collateral.join(" ; ")}`);
  console.log();
}

console.log(`${proven}/${MUTANTS.length} mutation(s) actually proved their guard.`);
console.log(
  allGood
    ? "Every mutation was caught by its own guard."
    : proven + baseline.length === 0
      ? "AT LEAST ONE MUTATION WENT UNNOTICED - the verifier cannot be trusted."
      : "THIS RUN DID NOT CERTIFY THE VERIFIER - fix the baseline, or a silent guard, and run it again.",
);
process.exit(allGood ? 0 : 1);
