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

/**
 * Whatever was thrown, as a bounded sentence.
 *
 * `String(e)` was the fallback, and for an object it reaches Object's default
 * stringification: a thrown `{ message: "InvalidArgument", code: 5 }` reached
 * the user as **"[object Object]"**, throwing away the message sitting inside
 * it. `undefined` reached them as the word "undefined".
 *
 * `formatValue` in `trace.ts` already refuses to do that, with the same
 * reasoning written next to it — "a line that occupies space and answers
 * nothing". This is the same rule on the path that reaches a PERSON rather
 * than a log.
 *
 * Every branch answers something a reader can act on or repeat to somebody who
 * can. The last one names the shape rather than pretending to a sentence,
 * because "the host raised something with no message in it" is a fact and
 * "[object Object]" is not.
 */
export function readable(e: unknown): string {
  if (e instanceof Error) return short(e.message);
  if (typeof e === "string") return short(e);
  if (e === null || e === undefined) return "the host raised nothing this pane can describe.";
  // An Office.js async failure is not always an `Error`: it is routinely a
  // plain object carrying `name`, `message` and `code`, and the message in it
  // is the whole point.
  if (typeof e === "object") {
    if ("message" in e && typeof e.message === "string") return short(e.message);
    try {
      // `JSON.stringify` answers `undefined` for a few shapes that reach here,
      // so the shape is named rather than falling through to a stringification
      // this whole function exists to avoid.
      return short(JSON.stringify(e) ?? "the host raised something with no message in it.");
    } catch {
      // Circular, which an Office error object can be.
      return "the host raised something this pane could not read.";
    }
  }
  // A thrown function is a caller's slip, and `String(fn)` prints its whole
  // source into the sentence. Named, not rendered — the same answer
  // `formatValue` gives it.
  if (typeof e === "function") return "the host raised a function, which is a bug in this add-in.";
  if (typeof e === "symbol") return short(e.toString());
  // Each primitive named, rather than one `String` over what is left. That is
  // the shape `formatValue` settled on for the same reason, and it is what the
  // `no-base-to-string` rule is asking for: `String` over a bare `unknown` is
  // exactly the call that put "[object Object]" on the screen.
  if (typeof e === "number" || typeof e === "boolean" || typeof e === "bigint") return short(String(e));
  return "the host raised something this pane could not read.";
}
