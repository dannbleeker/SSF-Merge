/**
 * Which source directories coverage measures, and which it does not.
 *
 * `src/core/trace.ts` names the hazard this exists for, and says it chose where
 * to live because of it: "the coverage config's `include` is a fixed list of
 * three globs — a new top-level directory would be measured by nothing, and an
 * uncounted module is how a threshold quietly stops meaning anything."
 *
 * A list here rather than inline in `vitest.config.ts` so a test can hold it
 * against what is actually on disk. Importing the config itself was the first
 * attempt and it does not work: pulling it into a test pulls it into the
 * TypeScript project service, and lint refuses the file for being in two
 * projects at once.
 */

/** Measured, and held to the thresholds. */
export const MEASURED = ["core", "host", "pane"];

/**
 * Not measured, with the reason, because a glob cannot carry one.
 *
 * Adding a directory to `src/` should be a decision somebody writes down rather
 * than a line they forget: the thresholds go on passing either way, which is
 * what makes the omission quiet.
 */
export const NOT_MEASURED = {
  office:
    "the Office.js calls themselves — they cannot run in the suite, and every decision they make lives in src/host",
};
