/**
 * Whether this module is the entry point.
 *
 * `import.meta.url.endsWith(process.argv[1].split("/").pop())` is the form
 * that reads correctly and never splits a BACKSLASHED path, so on Windows every
 * invocation prints nothing and exits 0 — which reads exactly like a pass. A
 * sibling project lost three tool CLIs to it for months. Compared as file URLs,
 * both platforms answer the same question.
 */
import { pathToFileURL } from "node:url";

export function isMain(metaUrl, argv = process.argv) {
  const entry = argv[1];
  if (!entry) return false;
  return metaUrl === pathToFileURL(entry).href;
}
