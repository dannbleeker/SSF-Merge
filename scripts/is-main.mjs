/**
 * Whether this module is the entry point.
 *
 * `import.meta.url.endsWith(process.argv[1].split("/").pop())` is the form
 * that reads correctly and never splits a BACKSLASHED path, so on Windows every
 * invocation prints nothing and exits 0 — which reads exactly like a pass. A
 * sibling project lost three tool CLIs to it for months. Compared as file URLs,
 * both platforms answer the same question.
 *
 * **Through the REAL path, because a symlink is the same failure again.** Node
 * gives `import.meta.url` as the resolved target and leaves `argv[1]` as the
 * link, so the two never matched and the CLI did nothing, silently, exiting 0.
 * `npm run` passes a direct path and is unaffected; `npm link`, a
 * `node_modules/.bin` shim and a global install all go through a link. Eight
 * scripts import this, `check-release.mjs` among them — a release pre-flight
 * that quietly declines to run is worse than one that fails.
 *
 * `realpathSync` throws for a path that is not there, which is not a reason to
 * lose the answer: an entry that cannot be resolved is compared as written,
 * which is what this did before.
 */
import { realpathSync } from "node:fs";
import { pathToFileURL } from "node:url";

export function isMain(metaUrl, argv = process.argv) {
  const entry = argv[1];
  if (!entry) return false;
  let resolved = entry;
  try {
    resolved = realpathSync(entry);
  } catch {
    /* not on disk, or not readable: compare what we were given */
  }
  return metaUrl === pathToFileURL(resolved).href;
}
