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
      return state.block ? null : "Choose the slides that repeat first.";
    case "preview":
      if (!state.block) return "Choose the slides that repeat first.";
      if (!state.rows) return "Attach your data first.";
      return null;
    case "merge": {
      if (!state.block) return "Choose the slides that repeat first.";
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
  switch (step) {
    case "template":
      return state.block
        ? { label: `Use slides ${state.block.from} to ${state.block.to}`, enabled: true }
        : { label: "Choose the slides that repeat", enabled: false };
    case "fields":
      return { label: "Attach data", enabled: reachable };
    case "preview":
      return state.previewing
        ? { label: "Put the template back", enabled: true }
        : { label: "Preview the first row", enabled: reachable };
    case "merge": {
      const n = state.block && state.rows ? slidesPerRecord(state.block) * state.rows : 0;
      return { label: n > 0 ? `Add ${n} slide${n === 1 ? "" : "s"}` : "Add slides", enabled: reachable && n > 0 };
    }
  }
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
