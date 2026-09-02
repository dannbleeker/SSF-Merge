/**
 * Wording that more than one layer has to get right the same way.
 *
 * There is exactly one thing in here so far, and it earned its place by being
 * spelled out by hand in seven different files. The pane's button said "Use
 * slides 2 to 2" for a one-slide block while the heading directly above it said
 * "Slide 2", because one of the seven asked a helper and the other six built the
 * phrase themselves.
 *
 * It lives in `core` because that is the only layer `pane`, `host` and `office`
 * can all import from, and it is deliberately NOT in `core/merge/text.ts`:
 * that file is the placeholder engine, and a sentence a user reads has no
 * business in it.
 */

/**
 * A slide range as a reader says it: "slide 3", or "slides 3 to 5".
 *
 * Lower case, because every caller so far puts it mid-sentence — "Use slide 3",
 * "The template block is slides 3 to 5", "none of slides 4 to 9 could be shown
 * to be this merge's". A caller that needs it as a sentence subject capitalises
 * it; `blockName` in the pane is the one that does.
 *
 * `from === to` is the whole point. A range with itself at both ends reads as
 * something the user mistyped rather than a sentence the product meant, and it
 * is the commonest block of all: one slide per row.
 */
export function slideRange(from: number, to: number): string {
  return from === to ? `slide ${from}` : `slides ${from} to ${to}`;
}
