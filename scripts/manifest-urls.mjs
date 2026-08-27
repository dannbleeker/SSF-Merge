#!/usr/bin/env node
/**
 * Every URL a production manifest points at, one per line.
 *
 * The release workflow fetches these before it creates anything. Nothing else
 * in this repo ever asks whether the host a manifest names is actually SERVING:
 * `checkManifest` asks whether a production manifest points at localhost, which
 * is a different question and passes cleanly for a manifest pointing at a
 * domain that 404s.
 *
 * That matters here more than it would elsewhere, because the release note tells
 * people to sideload this file and the pane is not IN it — the manifest is four
 * URLs into GitHub Pages. A renamed asset, a Pages deployment that never ran, a
 * DNS record that expired, and the add-in installs perfectly and shows a blank
 * ribbon button and an empty pane, with nothing anywhere saying why.
 *
 * The listing lives here and the FETCHING lives in the workflow, deliberately:
 * `urlsIn` is the tested part and a network call is not testable in this suite.
 * Run on the release job only, never on CI — a third-party outage must not block
 * a merge, which is the same reasoning that keeps Microsoft's validator out of
 * the `test` job.
 *
 *   node scripts/manifest-urls.mjs manifest-prod.xml
 */
import { readFileSync } from "node:fs";
import { urlsIn } from "./manifest-rules.mjs";
import { isMain } from "./is-main.mjs";

/**
 * The URLs worth fetching, out of the ones a manifest names.
 *
 * Only the ones this project SERVES. A manifest carries a support link and a
 * source-code link too, and a release must not be blocked because github.com
 * rate-limited a runner or a documentation page moved — those are somebody
 * else's uptime. Same-origin as the taskpane is the test, because that origin
 * is the one this repo deploys.
 *
 * @param {string[]} urls
 * @returns {string[]}
 */
export function fetchable(urls) {
  const pane = urls.find((u) => u.endsWith("taskpane.html"));
  if (!pane) return [];
  const origin = new URL(pane).origin;
  return [...new Set(urls.filter((u) => u.startsWith(`${origin}/`)))].sort();
}

/** @param {string[]} argv */
export function main(argv = process.argv.slice(2)) {
  const files = argv.length > 0 ? argv : ["manifest-prod.xml"];
  /** @type {Set<string>} */
  const urls = new Set();
  for (const file of files) for (const url of fetchable(urlsIn(readFileSync(file, "utf8")))) urls.add(url);
  for (const url of [...urls].sort()) console.log(url);
}

if (isMain(import.meta.url)) main();
