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
  /**
   * Rows the user has taken OUT, by index into `records.rows`.
   *
   * Excluded rather than included, so that the empty state means "merge
   * everything" — which is what a user who never opens the list wants, and
   * what a new paste must go back to. An included-list would default to empty
   * and merge nothing.
   */
  excluded?: number[];
  /** What is typed in the row search box. */
  rowSearch?: string;
  /** Whether the row list is open. Closed by default: 240 rows is not a screen. */
  rowsOpen?: boolean;
  /** Conditions the user set, keyed by SLIDE NUMBER — the numbering they can see. */
  conditions?: Record<number, string>;
  /** What the two slide-number boxes hold right now, as typed. */
  draft?: BlockDraft;
  /** What the user pasted into the data box, unparsed. */
  paste?: string;
  /** Something the pane has to say that is not a blocked step — a host answer. */
  notice?: string;
  /**
   * Whether this host can say which slides are selected (PowerPointApi 1.5).
   *
   * Undefined until the pane has asked, and the control is drawn only when it
   * is TRUE — an unasked host is treated the same as one that cannot, because
   * offering a shortcut that always fails is worse than not offering it.
   */
  canSelect?: boolean;
  /**
   * A host call this pane is waiting on.
   *
   * In the STATE, not on the button. The first version disabled the primary by
   * hand — `button.disabled = true` on a DOM node — and every later `draw()`
   * replaced that node with one `primary()` had re-enabled, because nothing
   * `primary()` can see said a run was out. Going back a step and forward again
   * during a 90-second template read handed the user a live "Add 720 slides"
   * over a merge already in flight.
   */
  running?: "inspect" | "merge" | "preview";
  /**
   * The host call in flight, for the pane to name while it waits.
   *
   * A merge is legitimately silent for up to two and a half minutes —
   * `BUDGET.file` allows ninety seconds to read the template and
   * `BUDGET.insert` sixty to hand the package over — and a frozen "Merging…"
   * for that long is indistinguishable from a wedged pane. Naming the call
   * costs nothing and is the difference between waiting and giving up.
   */
  inFlight?: string;
  /** The run record, shown so a user can copy it out. A task pane has no devtools. */
  log?: string;
  /**
   * Where a preview landed, in the numbering the thumbnail rail shows.
   *
   * The card names the slides rather than saying only "a preview is showing",
   * because a user who closes the pane mid-preview has no other way to find
   * out which slides to delete.
   */
  previewSlides?: Block;
  /**
   * What the last merge added, once it has.
   *
   * The merge screen is redrawn after a successful run, and without this it
   * redraws a live "Add 720 slides" beside a notice saying 720 slides were
   * added — one more press and there are 1440, in somebody's deck, from a
   * button that looks like the one they just pressed. Any edit clears it,
   * which is the only way back to an armed button.
   */
  added?: number;
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
   *
   * A `why` WITHOUT a block is a refusal. A `why` WITH one is advice the user
   * may press past — see the deck-size branch, which is the only thing here
   * measured against a number that can be stale.
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
  // Every refusal names what the USER typed, not just the rule. "Slides are
  // numbered from 1" is a true sentence that says nothing about the boxes in
  // front of them, and the manual promised numbers for all four cases while
  // two of them carried none.
  if (!Number.isInteger(a) || !Number.isInteger(b)) {
    const bad = !Number.isInteger(a) ? from : to;
    return { block: null, why: `Slide numbers are whole numbers, and "${bad}" is not one.` };
  }
  if (a < 1) return { block: null, why: `Slides are numbered from 1, so slide ${a} is not one.` };
  if (b < a) return { block: null, why: `The block ends before it starts: slide ${a} to ${b}.` };
  if (deckSize !== undefined && b > deckSize) {
    // ADVICE, not a refusal — the block comes back and the button stays live.
    // `deckSize` is a count taken when the pane opened, and a user who adds
    // slides to the deck and comes back would otherwise be told their block
    // does not exist, in a sentence stating a deck size that is no longer
    // true, with no way to correct it short of reopening the pane. The
    // authoritative check runs a moment later against ids the host listed just
    // now: `blockIds` refuses out of range, and `prepareBlock` refuses again
    // against the package that came back.
    return {
      block: { from: a, to: b },
      why: `Slide ${b} is past the end of the deck as this pane last counted it (${deckSize === 1 ? "1 slide" : `${deckSize} slides`}). If you have added slides since, go ahead.`,
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
      // Not the same as having no data. A user who unticked every row has
      // done something deliberate and needs telling what, not "attach data".
      if (includedCount(state) === 0) return "Every row is unticked, so there is nothing to merge.";
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

  // A host call is out. The label says which, and NOTHING is pressable — this
  // is the whole reason `running` is in the state rather than on the button.
  if (state.running === "inspect" && step === "template") {
    return { label: "Reading the slides…", enabled: false };
  }
  if (state.running === "merge" && step === "merge") {
    return { label: "Merging…", enabled: false };
  }
  if (state.running === "preview" && step === "preview") {
    // Inserting a preview is a real merge and can take a minute on this host.
    // Without this the button read "Preview the first row", greyed out, for the
    // whole of it — which is the state the user cannot tell from a pane that
    // has stopped responding. The other two long calls already named
    // themselves; this one was added with the step and missed.
    return { label: state.previewing ? "Removing…" : "Previewing…", enabled: false };
  }
  // Any other step, while something is out: it keeps its own words and loses
  // its press. Recursing with the flag cleared is deliberate — one place
  // decides what a step's button SAYS, and freezing must not fork it.
  if (state.running) return { ...primary({ ...state, running: undefined }, step), enabled: false };

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
      // "Remove", not "Put the template back". The template is never touched:
      // a preview is one row merged through the ORDINARY path and inserted, so
      // ending it deletes slides rather than restoring anything. The old label
      // described a design this project's own rejected list forbids.
      return state.previewing
        ? { label: "Remove the preview", enabled: true }
        : { label: "Preview the first row", enabled: reachable };
    case "merge": {
      // A run that already landed. Re-arming this button beside a notice
      // saying the slides were added is how a deck gets them twice.
      if (state.added !== undefined) {
        return { label: `Added ${state.added} slide${state.added === 1 ? "" : "s"}`, enabled: false };
      }
      // The INCLUDED rows, never the pasted ones. A user who has taken three
      // rows out and reads "Add 720 slides" has been told the wrong thing
      // about the button they are pressing.
      const n = block ? slidesPerRecord(block) * includedCount(state) : 0;
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

/**
 * The first row on its own, as a RecordSet the merge can run.
 *
 * A preview is not a separate rendering path — it is the REAL merge over one
 * record, inserted, looked at, and deleted again. That is what makes it worth
 * anything: what the user sees is produced by the code that will produce the
 * other 239 slides, not by a second implementation that can disagree with it.
 *
 * Named rather than inlined at the call site because "which row does a preview
 * show" is a decision, and an empty set has to answer something.
 */
export function firstRowOnly(records: RecordSet): RecordSet {
  return { columns: records.columns, rows: records.rows.slice(0, 1) };
}

/**
 * The row a preview shows, once rows can be taken out.
 *
 * The first row that will actually MERGE, not the first that was pasted. A
 * preview of a row the user has unticked is a preview of something nobody is
 * going to get — which is the same reason the preview runs the real merge in
 * the first place.
 */
export function firstIncludedRow(state: PaneState): RecordSet | undefined {
  const records = includedRecords(state);
  return records && records.rows.length > 0 ? firstRowOnly(records) : undefined;
}

/**
 * What a row is CALLED in the picker.
 *
 * The first column, because that is where a name, an id or an invoice number
 * lives in every table anyone pastes here — and if it is empty the row still
 * needs something to click, so it falls back to its position. Never the whole
 * row: at 320 pixels one column is what fits.
 */
export function rowLabel(records: RecordSet, index: number): string {
  const first = records.columns[0]?.name;
  const value = first === undefined ? "" : (records.rows[index]?.[first] ?? "");
  return value.trim() === "" ? `Row ${index + 1}` : value;
}

/**
 * The rows a search shows, as indices into `records.rows`.
 *
 * Matches across EVERY column, not just the labelled one. Someone looking for
 * "Aarhus" is looking for the row with Aarhus in it, and whether that happens
 * to be the column the label came from is not something they should have to
 * know. Case-insensitive; an empty query matches everything.
 */
export function visibleRows(records: RecordSet, query = ""): number[] {
  const q = query.trim().toLowerCase();
  const all = records.rows.map((_, i) => i);
  if (q === "") return all;
  return all.filter((i) => {
    const row = records.rows[i];
    if (!row) return false;
    return records.columns.some((c) => (row[c.name] ?? "").toLowerCase().includes(q));
  });
}

/** Whether a row is in the merge. Absent from `excluded` means yes. */
export function rowIncluded(state: PaneState, index: number): boolean {
  return !(state.excluded ?? []).includes(index);
}

/**
 * How many rows the merge will actually run.
 *
 * This is the number on the button and in the arithmetic — never
 * `records.rows.length`, which is how many were PASTED. A user who has taken
 * three rows out and reads "240 rows × 3 slides" has been told the wrong thing
 * about what they are pressing.
 */
export function includedCount(state: PaneState): number {
  // `rows` is the count and `records` is the data, and a state can carry the
  // first without the second — the pane knows how many rows it has before it
  // needs them. Reading only `records` made this answer ZERO for every such
  // state, which blocked the merge and emptied the button's number.
  const total = state.records ? state.records.rows.length : (state.rows ?? 0);
  // A Set, because a duplicate index in `excluded` would otherwise subtract
  // the same row twice and report fewer rows than will merge.
  const out = new Set((state.excluded ?? []).filter((i) => i >= 0 && i < total));
  return total - out.size;
}

/**
 * The records the merge runs on, with the excluded rows gone.
 *
 * A new RecordSet rather than a flag the engine has to honour: the engine's
 * whole contract is "one set of slides per row of what you give me", and
 * filtering at the boundary keeps it that way. Column order and types are the
 * ones the parse produced, because a filter removes rows and nothing else.
 */
export function includedRecords(state: PaneState): RecordSet | undefined {
  if (!state.records) return undefined;
  const rows = state.records.rows.filter((_, i) => rowIncluded(state, i));
  return { columns: state.records.columns, rows };
}
