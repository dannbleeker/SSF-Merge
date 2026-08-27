#!/usr/bin/env node
/**
 * Pre-flight for a release. Run by the workflow before it creates anything.
 *
 *   node scripts/check-release.mjs
 *
 * Exits non-zero with every problem named, so a bad release is refused before
 * a tag exists rather than yanked afterwards.
 */
import { readFileSync } from "node:fs";
import { RELEASE_ASSETS, releaseProblems } from "./release-assets.mjs";
import { isMain } from "./is-main.mjs";

export function main() {
  const problems = releaseProblems((name) => readFileSync(name, "utf8"));
  if (problems.length > 0) {
    for (const p of problems) console.error(`  ${p}`);
    throw new Error(`refusing to release: ${problems.length} problem(s)`);
  }
  console.log(`release assets ready: ${RELEASE_ASSETS.join(", ")}`);
}

if (isMain(import.meta.url)) main();
