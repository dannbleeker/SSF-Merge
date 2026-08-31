#!/usr/bin/env node
/**
 * Take the author's name out of a .pptx before it is committed.
 *
 *   node scripts/scrub-docprops.mjs test-kit/modern-chart.pptx [...]
 *
 * PowerPoint stamps whoever saved a deck into `docProps`, and this repository
 * is public: every deck it carries was authored by somebody, on purpose, so
 * that the engine is tested against parts it did not write itself. The test
 * kit's own README says a deck a human made by hand is the stronger test —
 * and the cost of that is a real name in the package, in a field nobody looks
 * at because no reader of a slide ever sees it.
 *
 * Four fields carry a person or an organisation. `dc:creator` and
 * `cp:lastModifiedBy` in `docProps/core.xml`, `Manager` and `Company` in
 * `docProps/app.xml`. Everything else in the package is left exactly as it
 * was, including the timestamps and the revision count: this is not an
 * anonymiser, it is four fields, and a script that quietly rewrote more of a
 * recording than it claimed would be worse than the problem.
 *
 * Idempotent, and safe to run on a deck that is already clean — it says so and
 * writes nothing, so it can be run over the whole set without churning files
 * whose bytes have not changed.
 *
 * `test/docprops.test.ts` is the standing check, and its failure message names
 * this script. The two exist together: a guard with no remedy is a guard the
 * next person works around.
 */
import { readFileSync, writeFileSync } from "node:fs";
import JSZip from "jszip";
import { isMain } from "./is-main.mjs";

/** The parts, and the fields in each, that Office writes a name into. */
const FIELDS = [
  { part: "docProps/core.xml", tags: ["dc:creator", "cp:lastModifiedBy"] },
  { part: "docProps/app.xml", tags: ["Manager", "Company"] },
];

/**
 * Empty one element, however it happens to be spelled.
 *
 * Both spellings are in this repository's own decks — python-pptx writes
 * `<dc:creator/>` and PowerPoint writes `<dc:creator></dc:creator>` — so a
 * scrubber that knew only one of them would report a deck clean because it
 * could not see the name in it.
 *
 * @param {string} xml
 * @param {string} tag
 * @returns {string}
 */
export function blankTag(xml, tag) {
  const escaped = tag.replace(":", "\\:");
  return xml
    .replace(new RegExp(`<${escaped}(\\s[^>]*)?/>`, "g"), `<${tag}></${tag}>`)
    .replace(new RegExp(`<${escaped}(\\s[^>]*)?>[\\s\\S]*?</${escaped}>`, "g"), `<${tag}></${tag}>`);
}

/**
 * What the named fields hold, for a report or a check.
 *
 * @param {string} xml
 * @param {string[]} tags
 * @returns {string[]}
 */
export function namesIn(xml, tags) {
  /** @type {string[]} */
  const held = [];
  for (const tag of tags) {
    const escaped = tag.replace(":", "\\:");
    const m = xml.match(new RegExp(`<${escaped}(?:\\s[^>]*)?>([\\s\\S]*?)</${escaped}>`));
    if (m && m[1].trim() !== "") held.push(`${tag}="${m[1].trim()}"`);
  }
  return held;
}

/** Scrub one package. Answers what it emptied, or an empty list for a clean deck. */
export async function scrub(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  /** @type {string[]} */
  const emptied = [];
  for (const { part, tags } of FIELDS) {
    const file = zip.file(part);
    if (!file) continue;
    const xml = await file.async("string");
    const held = namesIn(xml, tags);
    if (!held.length) continue;
    emptied.push(...held.map((h) => `${part} ${h}`));
    let next = xml;
    for (const tag of tags) next = blankTag(next, tag);
    // `createFolders: false` or jszip adds a `docProps/` directory entry the
    // deck never had. Harmless to a reader and still a change this script did
    // not intend: the claim it makes is four fields and nothing else.
    zip.file(part, next, { createFolders: false });
  }
  // Nothing to say and nothing to write: a re-zip of an unchanged deck still
  // moves every byte, and a file that changes for no reason is a file nobody
  // reads the diff of.
  if (!emptied.length) return { emptied, bytes: null };
  // DEFLATE, named rather than defaulted: jszip STORES unless told otherwise,
  // and the first version of this script tripled every deck it touched — 55 KB
  // to 193 KB, in files a public repository carries and a listing ships. A
  // .pptx is a zip, so `git` reports that as `Bin 54876 -> 193137 bytes` and
  // nothing else says a word.
  return { emptied, bytes: await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }) };
}

/**
 * Behind `isMain` because `test/docprops.test.ts` imports `namesIn` from here.
 * The field list has to be shared — a check and its remedy that each carry
 * their own copy will one day disagree about which fields matter — and a bare
 * CLI body runs on import, which took the test from failing to not collecting
 * at all.
 */
async function main() {
  const paths = process.argv.slice(2);
  if (paths.length === 0) {
    console.error("usage: node scripts/scrub-docprops.mjs <deck.pptx> [...]");
    process.exit(2);
  }
  for (const path of paths) {
    const { emptied, bytes } = await scrub(readFileSync(path));
    if (!bytes) {
      console.log(`${path}: already carries no name`);
      continue;
    }
    writeFileSync(path, bytes);
    console.log(`${path}: emptied ${emptied.join(", ")}`);
  }
}

if (isMain(import.meta.url)) await main();
