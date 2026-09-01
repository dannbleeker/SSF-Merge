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
import { chosenBlock, readBlockDraft, type BlockDraft, type PaneState } from "./steps.js";

/**
 * Everything that stops being true when the block moves.
 *
 * The block is the slides a merge repeats. Move it and the placeholders read
 * off the old slides are stale, the picture fields and the per-slide field
 * lists read with them are stale,
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
    // Read off the same slides in the same pass, and just as stale. Left
    // standing, `skippedRows` went on counting rows against the OLD block's
    // per-slide fields — dropping a row for a blank in a `{{Note}}` that is on
    // a slide this state no longer names, and saying so above the button.
    slideFields: undefined,
    conditions: undefined,
    // With the conditions, always: an anchor for conditions that are gone would
    // let the next keystroke keep conditions this one deliberately dropped.
    conditionsFor: undefined,
    changedSinceMerge: true,
    // The note names a token that was put on a slide in the OLD block. Left
    // standing it reports an insert into slides this state no longer names.
    fieldNote: undefined,
  };
}

/**
 * A new draft of the slide numbers — typed, or picked out of the deck.
 *
 * `blockMoved` is right for its own contract, and this is about WHEN it should
 * fire. It ran on every keystroke, so a user who set "slide 5 only when
 * Renewal", went back to check the slide numbers, retyped the same last slide
 * and walked forward again found the conditions gone — silently, with the merge
 * button quietly offering more slides than they had asked for.
 *
 * Both callers, and it was one for a while: the fix went into the typing path
 * and "use the slides I've selected" kept calling `blockMoved` directly, so
 * selecting the SAME slides still threw the conditions away. Two routes to one
 * question is how the first version of this defect was written.
 *
 * The conditions are keyed by SLIDE NUMBER, so what makes them stale is the
 * block naming different slides. Typing does not do that on its own: while a
 * box is empty the draft names no block at all, which is INCOMPLETE rather than
 * different, and a draft that resolves back to the same two numbers has not
 * moved anything. Only conditions are kept — the fields, the picture fields and
 * the per-slide lists were read off the slides themselves, and this pane has no
 * way to know the slides were not edited in between.
 */
export function blockDrafted(state: PaneState, draft: BlockDraft): PaneState {
  // The block the conditions are KEYED to, which outlives the empty box in the
  // middle of retyping a number. `chosenBlock` cannot answer for the second
  // keystroke: the first one already cleared the committed block, so from there
  // on the state names no block and every later keystroke would read as a move.
  const was = state.conditionsFor ?? chosenBlock(state);
  const now = readBlockDraft(draft, state.deckSize).block;
  const moved = was !== undefined && now !== null && (now.from !== was.from || now.to !== was.to);
  const next = { ...blockMoved(state), draft };
  if (moved || state.conditions === undefined || was === undefined) return next;
  return { ...next, conditions: state.conditions, conditionsFor: was };
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
