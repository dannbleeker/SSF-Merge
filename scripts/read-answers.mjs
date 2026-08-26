#!/usr/bin/env node
/**
 * Read an answer sheet the probe produced and say what it means.
 *
 * The snippet that runs inside PowerPoint collects raw observations and makes
 * no judgements: numbers, strings, and the errors it caught. Every reading
 * happens here, through `src/host/verdicts.ts`, which is covered by tests. A
 * probe that reasons inside the host is a probe whose conclusions nobody can
 * check.
 *
 *   node scripts/read-answers.mjs sheet.json
 *   node scripts/read-answers.mjs sheet.json --save
 */
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import {
  creationIdReading,
  insertionBlame,
  insertVerdict,
  offsetVerdict,
  Q3,
  Q4,
  substringVerdict,
  tagVerdict,
} from "../dist-lib/host/verdicts.js";

const file = process.argv[2];
if (!file) {
  console.error("usage: node scripts/read-answers.mjs <sheet.json> [--save]");
  process.exit(1);
}

const raw = readFileSync(file, "utf8");
// The snippet prints markers around the JSON so it can be copied out of a
// console. Tolerate a paste that still has them.
const body = raw.slice(raw.indexOf("{"), raw.lastIndexOf("}") + 1);
const sheet = JSON.parse(body);

const line = (label, value) => console.log(`  ${label.padEnd(22)} ${value}`);

console.log(`\nSSF Merge answer sheet — ${sheet.takenAt ?? "no timestamp"}`);
line("platform", sheet.platform ?? "unknown");
line("host version", sheet.host ?? "unknown");
line("PowerPointApi", (sheet.requirementSets ?? []).join(", ") || "none reported");
line("deck", `${sheet.deckAtStart} slides at start, ${sheet.deckAtEnd} at end`);

const fresh = insertVerdict({ ...sheet.insertFresh, expected: 2 });
const collision = insertVerdict({ ...sheet.insertCollision, expected: 2 });
const destTheme = sheet.insertFreshDestTheme
  ? insertVerdict({ ...sheet.insertFreshDestTheme, expected: 2 })
  : { verdict: "unknown", detail: "not asked — this sheet predates the arm", landed: 0 };
// The control's own deck is the presentation, so its size is what it inserts.
const self = sheet.insertSelf
  ? insertVerdict({ ...sheet.insertSelf, expected: sheet.deckAtStart })
  : { verdict: "unknown", detail: "not asked — this sheet predates the control arm", landed: 0 };

console.log("\n1. Does a cloned slide with a FRESH creation id insert?");
line("control: own deck", `${self.verdict} — ${self.detail}`);
line("fresh arm", `${fresh.verdict} — ${fresh.detail}`);
line("fresh, dest theme", `${destTheme.verdict} — ${destTheme.detail}`);
line("collision arm", `${collision.verdict} — ${collision.detail}`);
console.log(`\n  => ${insertionBlame(fresh.verdict, self.verdict)}`);
console.log(`  => ${creationIdReading(fresh, collision)}`);
if (fresh.verdict !== "yes" && destTheme.verdict === "yes") {
  console.log("  => and the THEME is the difference: the same package lands under UseDestinationTheme.");
}

console.log("\n2. Does a tag written into the PACKAGE read back through Office.js?");
const tag = tagVerdict({ ...(sheet.tagReadBack ?? {}), insertLanded: fresh.landed });
line("verdict", `${tag.verdict} — ${tag.detail}`);

const sub = sheet.substring ?? {};
console.log("\n3. Does a targeted substring write keep the formatting around it?");
if (sub.skipped) {
  line("verdict", `not asked — ${sub.skipped}`);
} else if (sub.error) {
  line("verdict", `threw — ${sub.error}`);
  line("threw at", sub.failedAt ?? "unknown — this sheet predates the step labels");
} else {
  const v = substringVerdict({
    before: sub.textBefore,
    after: sub.textAfterOne,
    want: Q3.want,
    boldAfter: sub.boldAfter,
  });
  line("verdict", `${v.verdict} — ${v.detail}`);
  line("text", JSON.stringify(sub.textAfterOne));
}

console.log("\n4. Do two writes queued in one batch interfere?");
if (sub.skipped) {
  line("verdict", `not asked — ${sub.skipped}`);
} else if (sub.error || !sub.twoWrites) {
  line("verdict", "not reached");
} else {
  const v = offsetVerdict(sub.twoWrites.after, Q4.independent, Q4.shifted);
  line("verdict", `${v.verdict} — ${v.detail}`);
  line("text", JSON.stringify(sub.twoWrites.after));
}

console.log("\n5. Does fill.setImage stretch or preserve aspect ratio?");
line(
  "verdict",
  "not asked here — no API reads it back. It is a look-at-the-slide question, and it gates image fields only.",
);

console.log(`\nclean-up: ${sheet.sweep ?? "not reported"}`);
if (sheet.deckAtEnd !== sheet.deckAtStart) {
  console.log(
    `  WARNING: the deck ended at ${sheet.deckAtEnd} and started at ${sheet.deckAtStart}. Check what was left behind.`,
  );
}

if (process.argv.includes("--save")) {
  mkdirSync("docs/host-answers", { recursive: true });
  const stamp = (sheet.takenAt ?? "unknown").replace(/[:.]/g, "-");
  const out = `docs/host-answers/${stamp}.json`;
  writeFileSync(out, `${JSON.stringify(sheet, null, 2)}\n`);
  console.log(`\nsaved ${out}`);
}
console.log();
