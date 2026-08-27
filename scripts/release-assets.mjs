/**
 * What a release ships, and the checks that stop it shipping the wrong thing.
 *
 * A release of this repo is small: the pane is served from Pages, so the only
 * thing a user downloads is a MANIFEST. That makes the whole risk surface one
 * question — is the file attached to the release the one the documentation told
 * them to download, pointing at production?
 *
 * Both halves of that have failed on a sibling project, and neither was caught
 * by anything for twelve days: a release that shipped the DEV manifests, while
 * the README pointed at a `manifest-prod.xml` that was not in the release at
 * all. Every gate in that repo read the working tree; a user downloads the
 * release, and the two had diverged.
 */
import { readFileSync } from "node:fs";
import { checkManifest, isProd } from "./manifest-rules.mjs";

/** The files a release attaches. Production only — see `releaseProblems`. */
export const RELEASE_ASSETS = ["manifest-prod.xml", "manifest-prod.json"];

/**
 * Every manifest the documentation tells a reader to download.
 *
 * Read out of the PROSE rather than listed here, because the failure this
 * guards against is the docs and the release disagreeing — and a second
 * hand-written list is a third thing that can disagree with both. Matches a
 * manifest filename wherever it appears in backticks or in bold.
 */
export function assetsPromisedByDocs(docs = ["docs/MANUAL.md", "README.md"]) {
  const promised = new Set();
  for (const path of docs) {
    const text = readFileSync(path, "utf8");
    for (const m of text.matchAll(/manifest(?:-prod)?\.(?:xml|json)/g)) promised.add(m[0]);
  }
  return [...promised].sort();
}

/**
 * What is wrong with this release, as sentences. Empty means nothing is.
 *
 * `read` is injected so the caller decides where the bytes come from — the
 * working tree here, and in the workflow the very files about to be uploaded.
 * Checking the tree and shipping something else is the whole failure mode.
 */
export function releaseProblems(read, assets = RELEASE_ASSETS, promised = assetsPromisedByDocs()) {
  const out = [];

  for (const name of assets) {
    if (!isProd(name)) {
      // A release that ships a dev manifest points every installer at
      // localhost, where nothing is listening.
      out.push(`${name} is not a production manifest, and a release must not ship one`);
      continue;
    }
    let text;
    try {
      text = read(name);
    } catch {
      out.push(`${name} is named as a release asset and is not there`);
      continue;
    }
    out.push(...checkManifest(text, name));
  }

  // The other direction, and the one that went unnoticed for twelve days: the
  // documentation naming a file the release does not carry.
  for (const name of promised) {
    if (!isProd(name)) continue; // a dev manifest is for contributors, not installers
    if (!assets.includes(name)) {
      out.push(`the documentation tells people to download ${name} and the release does not attach it`);
    }
  }
  return out;
}
