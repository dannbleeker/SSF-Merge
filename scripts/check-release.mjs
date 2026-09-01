#!/usr/bin/env node
/**
 * Pre-flight for a release. Run by the workflow before it creates anything.
 *
 *   node scripts/check-release.mjs
 *   RELEASE_VERSION=0.4.0 node scripts/check-release.mjs
 *
 * With a version it also asks whether this repo agrees that that is the version
 * — the workflow's dispatch box is free text.
 *
 * Exits non-zero with every problem named, so a bad release is refused before
 * a tag exists rather than yanked afterwards.
 */
import { readFileSync } from "node:fs";
import { RELEASE_ASSETS, releaseProblems } from "./release-assets.mjs";
import { isMain } from "./is-main.mjs";

export function main() {
  // The version the workflow was dispatched with, when there is one. Checked
  // against `package.json` and the changelog, because the dispatch box is free
  // text and nothing had ever looked at it.
  const version = process.env.RELEASE_VERSION ?? "";
  const problems = releaseProblems((name) => readFileSync(name, "utf8"), undefined, undefined, version);
  if (problems.length > 0) {
    for (const p of problems) console.error(`  ${p}`);
    throw new Error(`refusing to release: ${problems.length} problem(s)`);
  }
  console.log(`release assets ready: ${RELEASE_ASSETS.join(", ")}`);
}

if (isMain(import.meta.url)) main();
