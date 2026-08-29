/**
 * The state changes that touch more than one field at a time.
 *
 * A pane update that sets one field is readable where it happens. One that
 * clears six is a RULE — "this is what stops being true when the block moves" —
 * and a rule written out at each of its call sites is a rule that drifts.
 *
 * It had. Four places moved the block, and no two cleared the same fields:
 * typing in the slide-number boxes cleared `fieldNote`, picking a selection did
 * not, a failed read cleared neither `conditions` nor `added`, and none of them
 * cleared `imageFields`. The pane also once cleared `block` while leaving
 * `previewing` set, which left step 4 telling the user to choose slides above a
 * working "Remove the preview" button.
 *
 * Pure, so each rule can be checked without a DOM.
 */
import type { PaneState } from "./steps.js";

/**
 * Everything that stops being true when the block moves.
 *
 * The block is the slides a merge repeats. Move it and the placeholders read
 * off the old slides are stale, the picture fields read with them are stale,
 * the conditions are stale because they are keyed by SLIDE NUMBER — "slide 5
 * only when Renewal" is about the fifth slide of the deck, and a block starting
 * one slide later would apply it to a different slide — and a finished run's
 * disarmed button is about a merge this state no longer describes.
 *
 * `added` is deliberately NOT cleared. It is what the last run put IN THE
 * DECK, and moving the block does not take those slides out — clearing it took
 * the undo card off the screen while they were still there.
 *
 * `notice` is deliberately NOT here. Two of the callers set one of their own,
 * and clearing it inside this would make the order of two statements decide
 * whether the user sees the sentence.
 */
export function blockMoved(state: PaneState): PaneState {
  return {
    ...state,
    block: undefined,
    fields: [],
    imageFields: [],
    conditions: undefined,
    changedSinceMerge: true,
    // The note names a token that was put on a slide in the OLD block. Left
    // standing it reports an insert into slides this state no longer names.
    fieldNote: undefined,
  };
}

/**
 * Everything that stops being true when new data arrives.
 *
 * The row filter goes with the data it was about: row 7 of the old paste is not
 * row 7 of the new one, and carrying an exclusion across would take out a row
 * the user never looked at. Conditions are NOT cleared — a condition is about
 * the template, and a column the new data lacks is reported rather than quietly
 * dropped. Neither is `added`, for the reason `blockMoved` gives.
 */
export function dataChanged(state: PaneState): PaneState {
  return {
    ...state,
    changedSinceMerge: true,
    // A new paste can have different columns, so a note about `{{Region}}`
    // being placed may now be about a column nothing will fill.
    fieldNote: undefined,
    excluded: undefined,
    rowSearch: undefined,
  };
}
