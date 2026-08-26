/**
 * The four steps, and what may happen on each.
 *
 * Pure, because "can the user press Merge yet" is the question the pane is for
 * and the one most likely to be got wrong: it depends on a template block, on
 * data, and on whether every placeholder has a column behind it. Written inline
 * in a click handler it would be untestable and would drift from what the
 * screen says. Here it is a function of the state and the suite checks it.
 *
 * The step NUMBER is deliberately kept and shown ("Step 2 of 4"). It states how
 * much is left, which is the question a first-time user actually has, and it is
 * cheap to render. That was settled with the owner when the layout was.
 */

import { parseDelimited, toRecordSet } from "../core/data/recordset.js";
import type { RecordSet } from "../core/data/recordset.js";

export type StepId = "template" | "fields" | "preview" | "merge";

/** In order. The pane never renders these in any other sequence. */
export const STEPS: readonly StepId[] = ["template", "fields", "preview", "merge"] as const;

export const STEP_TITLE: Record<StepId, string> = {
  template: "Template",
  fields: "Fields",
  preview: "Preview",
  merge: "Merge",
};

/** A run of contiguous slides, in the numbering the THUMBNAIL RAIL shows. */
export interface Block {
  /** First slide of the template, 1-based, as the user sees it. */
  from: number;
  /** Last slide of the template, 1-based, inclusive. */
  to: number;
}

export interface PaneState {
  /** The template block, once the user has chosen one. */
  block?: Block;
  /** Placeholders found in the block, in the order they appear. */
  fields: string[];
  /** Column names in the data the user attached, if any. */
  columns?: string[];
  /** Rows in that data. */
  rows?: number;
  /** True while a preview is on the slide and has not been put back. */
  previewing: boolean;
  /** Slides in the deck right now, so the pane can say where the new ones land. */
  deckSize?: number;
  /** The parsed data, once attached. `rows` is its length, kept for the labels. */
  records?: RecordSet;
  /** Conditions the user set, keyed by SLIDE NUMBER — the numbering they can see. */
  conditions?: Record<number, string>;
  /** What the two slide-number boxes hold right now, as typed. */
  draft?: BlockDraft;
  /** What the user pasted into the data box, unparsed. */
  paste?: string;
  /** Something the pane has to say that is not a blocked step — a host answer. */
  notice?: string;
}

/**
 * The two slide-number boxes, as STRINGS.
 *
 * Not numbers, because "" and "-" and "0" are all states a box passes through
 * on the way to a valid one, and a number type has no way to hold them: an
 * `<input type="number">` reports "" for both an empty box and one holding
 * `"--"`, and coercing on the way in turns a half-typed entry into a 0 the pane
 * then complains about. The read is a separate function so the complaint can be
 * withheld until there is something to complain about.
 */
export interface BlockDraft {
  from: string;
  to: string;
}

export const EMPTY_DRAFT: BlockDraft = { from: "", to: "" };

export interface BlockRead {
  /** The block, once both boxes hold numbers that make sense together. */
  block: Block | null;
  /**
   * What is wrong, or null.
   *
   * Null while either box is still EMPTY. A form that turns red on the first
   * keystroke is a form that is wrong more often than the user is, and the
   * boxes are filled one at a time — so "4" and "" is a half-typed entry, not a
   * mistake.
   */
  why: string | null;
}

/**
 * Read the two boxes.
 *
 * `deckSize` is optional because the pane may not have counted the deck yet,
 * and a block that runs past the end is better caught here — where the sentence
 * can name both numbers — than by `prepareBlock` after a whole template read.
 * The wording matches `prepareBlock`'s deliberately: the same mistake should
 * not have two different explanations depending on which layer noticed it.
 */
export function readBlockDraft(draft: BlockDraft, deckSize?: number): BlockRead {
  const from = draft.from.trim();
  const to = draft.to.trim();
  if (from === "" || to === "") return { block: null, why: null };

  const a = Number(from);
  const b = Number(to);
  if (!Number.isInteger(a) || !Number.isInteger(b)) {
    return { block: null, why: "The template block has to be whole slide numbers." };
  }
  if (a < 1) return { block: null, why: "Slides are numbered from 1." };
  if (b < a) return { block: null, why: `The block ends before it starts: slide ${a} to ${b}.` };
  if (deckSize !== undefined && b > deckSize) {
    return {
      block: null,
      why: `The block ends at slide ${b} and the deck has ${deckSize === 1 ? "1 slide" : `${deckSize} slides`}.`,
    };
  }
  return { block: { from: a, to: b }, why: null };
}

/**
 * The block the pane should act on.
 *
 * The DRAFT wins over the committed block, because the draft is what the user
 * is looking at: a button that still says "Use slides 4 to 6" while the boxes
 * read 7 and 9 is a button that does something other than what it says.
 */
export function chosenBlock(state: PaneState): Block | undefined {
  const drafted = state.draft ? readBlockDraft(state.draft, state.deckSize).block : null;
  return drafted ?? state.block;
}

export interface DataRead {
  records: RecordSet | null;
  columns: string[];
  rows: number;
  /** What is wrong, or null. Null for an empty box — nothing pasted yet. */
  why: string | null;
}

/**
 * Read whatever was pasted into the data box.
 *
 * ONE function, so the chips, the row count, the button's number and the
 * records the merge actually runs on all come from the same parse. The
 * alternative — parsing in the input handler for the labels and again in the
 * merge for the data — is two parses that can disagree, and the one that
 * disagrees is the one the user never sees.
 *
 * The delimiter is left to `parseDelimited`, which sniffs tab first: the
 * commonest input by far is a range copied out of Excel, and that arrives
 * tab-separated.
 */
export function readPastedTable(text: string): DataRead {
  const empty: DataRead = { records: null, columns: [], rows: 0, why: null };
  if (text.trim() === "") return empty;
  const records = toRecordSet(parseDelimited(text));
  if (records.rows.length === 0) {
    return { ...empty, why: "That is a header row with no data under it." };
  }
  return {
    records,
    columns: records.columns.map((c) => c.name),
    rows: records.rows.length,
    why: null,
  };
}

export const EMPTY: PaneState = { fields: [], previewing: false };

/** How many slides one record produces. */
export function slidesPerRecord(block: Block): number {
  return block.to - block.from + 1;
}

/**
 * Placeholders with no column behind them.
 *
 * Named rather than inlined because it is the one thing that blocks the merge
 * for a reason the user can fix, and the pane has to say WHICH.
 */
export function unmatchedFields(state: PaneState): string[] {
  if (!state.columns) return [];
  const have = new Set(state.columns);
  return state.fields.filter((f) => !have.has(f));
}

export type Status = "done" | "current" | "waiting";

/**
 * Why a step cannot be reached yet, in words the pane shows as-is.
 *
 * Null means the step is reachable. Every sentence names the thing the user has
 * to do, because "Not available" is a dead end and this pane has four places it
 * could have said that.
 */
export function blockedReason(state: PaneState, step: StepId): string | null {
  switch (step) {
    case "template":
      return null;
    case "fields":
      return chosenBlock(state) ? null : "Choose the slides that repeat first.";
    case "preview":
      if (!chosenBlock(state)) return "Choose the slides that repeat first.";
      if (!state.rows) return "Attach your data first.";
      return null;
    case "merge": {
      if (!chosenBlock(state)) return "Choose the slides that repeat first.";
      if (!state.rows) return "Attach your data first.";
      const missing = unmatchedFields(state);
      if (missing.length > 0) {
        // Name them. A count alone sends the user back through every slide.
        return `No column for ${missing.join(", ")}. Rename the column or the placeholder.`;
      }
      if (state.previewing) return "End the preview before merging.";
      return null;
    }
  }
}

export function statusOf(state: PaneState, step: StepId, current: StepId): Status {
  if (step === current) return "current";
  return blockedReason(state, step) === null && STEPS.indexOf(step) < STEPS.indexOf(current) ? "done" : "waiting";
}

/** The one filled button on a screen: what it says, and whether it can be pressed. */
export interface Primary {
  label: string;
  enabled: boolean;
}

/**
 * Exactly one primary per step, always naming what it does WITH THE NUMBER IN
 * IT. "Merge" is a promise; "Add 720 slides" is a statement the user can check
 * against the deck they are looking at.
 */
export function primary(state: PaneState, step: StepId): Primary {
  const reachable = blockedReason(state, step) === null;
  const block = chosenBlock(state);
  switch (step) {
    case "template":
      return block
        ? { label: `Use slides ${block.from} to ${block.to}`, enabled: true }
        : { label: "Choose the slides that repeat", enabled: false };
    case "fields":
      // "Attach data" is what the step is FOR, so it stays the label until
      // there is data; once there is, the button states what it will carry
      // forward. A button that says "Attach data" after the data is attached
      // reads as a step that did not take.
      return state.rows
        ? { label: `Use ${state.rows} row${state.rows === 1 ? "" : "s"}`, enabled: reachable }
        : { label: "Attach data", enabled: false };
    case "preview":
      // NOT "Preview the first row". Nothing writes a preview to the slide
      // yet, and a button naming something that does not happen is worse than
      // one that names the step it does. The previewing branch is untouched
      // because putting a preview BACK is the half that is built.
      return state.previewing
        ? { label: "Put the template back", enabled: true }
        : { label: "Continue to merge", enabled: reachable };
    case "merge": {
      const n = block && state.rows ? slidesPerRecord(block) * state.rows : 0;
      return { label: n > 0 ? `Add ${n} slide${n === 1 ? "" : "s"}` : "Add slides", enabled: reachable && n > 0 };
    }
  }
}

/**
 * The step after this one, or null at the end.
 *
 * Null rather than a wrap, and null rather than the first step, because the
 * first version of this was `order[order.indexOf(from) + 1]` in `main.ts` with
 * no bound: `indexOf` answers -1 for anything that is not a step, so a stray
 * `data-action` sent the user back to step 1 with their block and their data
 * still in state — a wizard that resets itself and looks like it lost the lot.
 */
export function nextStep(from: StepId): StepId | null {
  const i = STEPS.indexOf(from);
  if (i < 0) return null;
  return STEPS[i + 1] ?? null;
}

/**
 * Which element in a view is allowed to be orange.
 *
 * The layout was approved on an ORANGE BUDGET of one per view: normally the
 * tick above the heading, and when something is temporarily untrue on the slide
 * the orange moves to that warning and the tick goes away. Two oranges in one
 * glance and neither means anything.
 *
 * A single function rather than a condition at each call site, because the
 * first version WAS conditions at each call site and it broke immediately: the
 * fields step drew the tick and an orange-bordered chip for an unmatched
 * placeholder, two oranges, and nothing but a screenshot could see it. A budget
 * enforced in one place can be tested; one enforced in three cannot.
 *
 * Precedence is preview first, because a preview is a state the user must undo
 * and an unmatched field is one they can leave sitting there.
 */
export type OrangeHolder = "tick" | "preview" | "unmatched";

export function orangeHolder(state: PaneState, step: StepId): OrangeHolder {
  if (state.previewing) return "preview";
  if (step === "fields" && unmatchedFields(state).length > 0) return "unmatched";
  return "tick";
}
