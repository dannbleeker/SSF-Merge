#!/usr/bin/env node
/**
 * Write the four manifests from the one definition.
 *
 * Committed rather than built on demand, because a manifest is a file a PERSON
 * sideloads: it has to be downloadable from the repo at a stable path without a
 * toolchain. `test/manifest.test.ts` fails when a committed file stops matching
 * what `manifest-source.mjs` produces, so "committed" and "generated" cannot
 * come apart the way four hand-written files would.
 *
 *   node scripts/build-manifests.mjs
 */
import { writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { allManifests } from "./manifest-source.mjs";
import { checkManifest } from "./manifest-rules.mjs";
import { isMain } from "./is-main.mjs";

export function main() {
  const root = dirname(dirname(fileURLToPath(import.meta.url)));
  const all = allManifests();
  const problems = [];
  for (const [name, text] of Object.entries(all)) problems.push(...checkManifest(text, name));
  if (problems.length > 0) {
    // Refused rather than written. A generator that emits a manifest its own
    // rules reject has produced a file somebody will sideload.
    for (const p of problems) console.error(`  ${p}`);
    throw new Error(`refusing to write ${problems.length} manifest problem(s)`);
  }
  for (const [name, text] of Object.entries(all)) writeFileSync(join(root, name), text);
  console.log(`manifests: ${Object.keys(all).join(", ")}`);
}

if (isMain(import.meta.url)) main();
