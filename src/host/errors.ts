/**
 * A raise as a sentence, bounded.
 *
 * Two copies of `readable` existed, in `src/office/merge.ts` and
 * `src/pane/main.ts`, and neither capped anything. That matters here more than
 * it usually would: Office echoes an argument back into `debugInfo`, and the
 * argument to `insertSlidesFromBase64` is the entire merged deck as base64.
 * A 400-slide merge is tens of megabytes, and the path from that string to the
 * screen was `err.message` → `state.notice` → a DOM text node, uncapped at
 * every step.
 *
 * So a failed insert could put the whole package into the pane as text, with
 * the sentence that explains the failure at the front and nothing after it
 * readable. The sibling project has the same defect recorded from the other
 * end: a diagnosis of theirs sat behind about 100 KB of base64 for the same
 * reason.
 */

/** Longest error text that reaches a user. Enough for a real sentence. */
export const ERROR_CHARS = 400;

/** `s`, cut to `max` with what was dropped counted rather than merely elided. */
export function short(s: string, max = ERROR_CHARS): string {
  // The count, not a bare ellipsis: "…" alone leaves a reader unable to tell a
  // truncated sentence from one that ended oddly, and unable to say whether
  // what is missing was a paragraph or a megabyte.
  return s.length <= max ? s : `${s.slice(0, max)}… (${s.length - max} more characters)`;
}

/** Whatever was thrown, as a bounded sentence. */
export function readable(e: unknown): string {
  return short(e instanceof Error ? e.message : String(e));
}
