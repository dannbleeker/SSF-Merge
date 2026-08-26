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
import { creationIdReading, insertVerdict, offsetVerdict, substringVerdict } from "../dist-lib/host/verdicts.js";

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

console.log("\n1. Does a cloned slide with a FRESH creation id insert?");
line("fresh arm", `${fresh.verdict} — ${fresh.detail}`);
line("collision arm", `${collision.verdict} — ${collision.detail}`);
console.log(`\n  => ${creationIdReading(fresh, collision)}`);

console.log("\n2. Does a tag written into the PACKAGE read back through Office.js?");
const tag = sheet.tagReadBack ?? {};
if (tag.error) line("verdict", `threw — ${tag.error}`);
else if (tag.value === "probe-run")
  line("verdict", "yes — the whole metadata scheme works, and no tag write is needed in the host");
else if (tag.value === undefined)
  line("verdict", "NO — the host did not find the tag. The metadata scheme needs rethinking before the pane is built.");
else line("verdict", `unexpected value ${JSON.stringify(tag.value)}`);

const sub = sheet.substring ?? {};
console.log("\n3. Does a targeted substring write keep the formatting around it?");
if (sub.error) {
  line("verdict", `threw — ${sub.error}`);
} else {
  const v = substringVerdict({
    before: sub.textBefore,
    after: sub.textAfterOne,
    want: "Hello Ada here and AAA-BBB",
    boldAfter: sub.boldAfter,
  });
  line("verdict", `${v.verdict} — ${v.detail}`);
  line("text", JSON.stringify(sub.textAfterOne));
}

console.log("\n4. Do two writes queued in one batch interfere?");
if (sub.error || !sub.twoWrites) {
  line("verdict", "not reached");
} else {
  const v = offsetVerdict(sub.twoWrites.after, "Hello Ada here and 1-2", "Hello Ada here and 1-2".replace("2", "BBB"));
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
