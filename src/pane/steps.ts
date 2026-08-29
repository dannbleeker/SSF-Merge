/**
 * The five steps, and what may happen on each.
 *
 * Pure, because "can the user press Merge yet" is the question the pane is for
 * and the one most likely to be got wrong: it depends on a template block, on
 * data, and on whether every placeholder has a column behind it. Written inline
 * in a click handler it would be untestable and would drift from what the
 * screen says. Here it is a function of the state and the suite checks it.
 *
 * The step NUMBER is deliberately kept and shown ("Step 2 of 5"). It states how
 * much is left, which is the question a first-time user actually has, and it is
 * cheap to render. That was settled with the owner when the layout was.
 */

import { canBeField } from "../core/merge/text.js";
import { slideApplies } from "../core/merge/plan.js";
import { baseName } from "../core/merge/images.js";
import { imageNamesIn, parseDelimited, toRecordSet } from "../core/data/recordset.js";
import type { RecordSet } from "../core/data/recordset.js";

/**
 * The order, and why DATA comes before FIELDS.
 *
 * A field is `{{Column}}` and the column names live in the user's data, so
 * there is nothing to insert until the data is attached. The first version put
 * the fields step second and asked the user to type placeholders they had no
 * way to name yet — reported from a first real run on a fresh deck, where the
 * refusal read as the add-in being broken.
 *
 * So: mark the slides that repeat, paste the rows, put the fields on the
 * slides, look at one, merge.
 */
export type StepId = "template" | "data" | "fields" | "preview" | "merge";

/** In order. The pane never renders these in any other sequence. */
export const STEPS: readonly StepId[] = ["template", "data", "fields", "preview", "merge"] as const;

export const STEP_TITLE: Record<StepId, string> = {
  template: "Template",
  data: "Data",
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
  /**
   * Of those, the ones written as a PICTURE — `{{Photo|image}}` and its two
   * siblings — read off the slides by the engine.
   *
   * The pane used to decide what a picture was from the DATA's detected types
   * alone, and the engine decides from the FIELD's format. See
   * `pictureColumns`.
   */
  imageFields?: string[];
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
  /**
   * The picture files the user picked, by the name they were picked under.
   *
   * Bytes, in the pane, never leaving it. The alternative was a URL per row,
   * which a task pane cannot fetch anyway — it is a sandboxed cross-origin
   * iframe and most image hosts send no CORS headers — and which would mean the
   * add-in making a request out of somebody's data for every row of it.
   *
   * A Map rather than a record because a file name is not a safe object key,
   * and because the engine takes one: this is handed to `runPlan` as it stands.
   */
  images?: Map<string, Uint8Array>;
  /** Whether the row list is open. Closed by default: 240 rows is not a screen. */
  rowsOpen?: boolean;
  /**
   * Whether the condition list is open. Closed by default, like the rows.
   *
   * Most merges want every slide for every row, and a user who never opens this
   * should pay one line for it. Shut, that line also STATES the current answer
   * ("every slide, every row"), so the feature is discoverable without being in
   * the way — a `<details>` whose summary says nothing is a control nobody
   * finds.
   */
  conditionsOpen?: boolean;
  /** Conditions the user set, keyed by SLIDE NUMBER — the numbering they can see. */
  conditions?: Record<number, string>;
  /** What the two slide-number boxes hold right now, as typed. */
  draft?: BlockDraft;
  /** What the user pasted into the data box, unparsed. */
  paste?: string;
  /** Something the pane has to say that is not a blocked step — a host answer. */
  notice?: string;
  /**
   * What the last press of an Insert button did, in words.
   *
   * Its own field rather than `notice`, because the two survive different
   * things: a notice is cleared by the next edit, and this has to stay on
   * screen while the user leaves the pane, clicks into a text box on the slide
   * and comes back. It is also the ONLY report the clipboard fallback has —
   * where the insert lands visibly on the slide, a copy lands nowhere the user
   * can see.
   */
  fieldNote?: string;
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
  running?: "inspect" | "merge" | "preview" | "undo";
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
  /**
   * How many slides the deck held BEFORE that merge.
   *
   * The undo card is a positional offer, and position means nothing without
   * this: `added` alone says how many slides to take, not which. With the
   * deck's size now it says both, and it says them through `sweepPlan`, so the
   * sentence on the card is the range the button removes.
   *
   * Set wherever `added` is — after a run, and from the crash crumb, which has
   * carried it since it was written.
   */
  deckAtStart?: number;
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

  // The TEXT has to look like a decimal number before `Number` is asked what it
  // is worth. `Number` reads far more than anybody types into a slide box: it
  // takes `0x10` as sixteen, `0b11` as three, `0o17` as fifteen and `1e2` as a
  // hundred.
  //
  // Both halves of that were wrong here. `0b11` was ACCEPTED, so the pane
  // quietly merged slides 3 to 9 for somebody who had typed neither number. And
  // the refusals named a cause they had invented: `0x10` produced "The block
  // ends before it starts: slide 16 to 9", about a slide 16 that appears
  // nowhere on the user's screen and in nothing they typed.
  //
  // A fractional part of zeros stays admitted, because `4.0` IS a whole number
  // and refusing it with "4.0 is not one" would be a false sentence.
  //
  // Third time `Number()` has been too wide in this codebase — it took `0x10`
  // out of a data cell as sixteen, and admitted an unreadable grouping — so the
  // rule is the same one: ask the shape first.
  const DECIMAL = /^[+-]?\d+(?:\.\d+)?$/;
  const a = DECIMAL.test(from) ? Number(from) : Number.NaN;
  const b = DECIMAL.test(to) ? Number(to) : Number.NaN;
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
 * The slide numbers in the chosen block, in the numbering the user sees.
 *
 * Conditions are keyed by SLIDE NUMBER because that is what the user picked in
 * step 1 and what the thumbnail rail shows. `prepareBlock` reads them back the
 * same way (`req.conditions?.[req.from + i]`), so the pane and the engine agree
 * on what "slide 5" means without either converting.
 */
/**
 * How many slides this merge will actually add.
 *
 * Not slides-per-record times rows. That product ignores the conditions, so a
 * block with one conditional slide promised more slides than the plan builds —
 * in the sentence directly above the button, along with a deck size that was
 * wrong by the same amount.
 *
 * Counted with `slideApplies`, which is the rule `buildPlan` itself applies, so
 * the promise and the plan cannot answer differently.
 */
export function plannedSlides(state: PaneState): number {
  const block = chosenBlock(state);
  if (!block) return 0;
  const records = includedRecords(state);
  if (!records) return 0;
  const columns = new Set(records.columns.map((c) => c.name));
  let n = 0;
  for (const row of records.rows) {
    for (const slide of blockSlides(state)) {
      const condition = conditionFor(state, slide);
      if (slideApplies(condition === "" ? {} : { condition }, row, columns)) n++;
    }
  }
  return n;
}

/**
 * How many slides this merge will add — the ONE answer, for every reader.
 *
 * `plannedSlides` was written because the product of slides-per-record and rows
 * ignores conditions, so the forecast above the button promised more slides
 * than the plan builds. The BUTTON kept the product: with one conditional slide
 * and three rows the card read "4 slides added after slide 20" and the button
 * beside it read "Add 6 slides". Two numbers on one screen, and the wrong one
 * on the thing being pressed.
 *
 * The suite already recorded that the two answers differ — `plannedSlides` has
 * an assertion that it is NOT the product, so that its own test cannot pass
 * against the bug. Nothing asked whether anything still used the product.
 *
 * The product survives here for the one state where it is the only answer
 * available: rows counted, data not yet in hand. A condition cannot be
 * evaluated without a row, and `includedCount` works from `state.rows` for
 * exactly that state — reading only `records` made it answer zero and emptied
 * the button's number, which is a defect this file already carries a comment
 * about.
 */
export function slidesToAdd(state: PaneState): number {
  const block = chosenBlock(state);
  if (!block) return 0;
  if (state.records) return plannedSlides(state);
  return slidesPerRecord(block) * includedCount(state);
}

export function blockSlides(state: PaneState): number[] {
  const block = chosenBlock(state);
  if (!block) return [];
  const out: number[] = [];
  for (let n = block.from; n <= block.to; n++) out.push(n);
  return out;
}

/** The column a slide is conditional on, or "" for "always". */
export function conditionFor(state: PaneState, slide: number): string {
  return state.conditions?.[slide] ?? "";
}

/**
 * The conditions with one slide's answer changed.
 *
 * Returns `undefined` when nothing is left, so the state carries no empty
 * object: `state.conditions` being absent and being `{}` mean the same thing to
 * every reader, and two spellings of one meaning is how a comparison starts
 * disagreeing with itself.
 *
 * An empty column is a DELETE rather than a stored blank. `prepareBlock` tests
 * the value for truthiness, so a stored `""` would behave as "always" — right
 * by accident, and it would still be reported as a condition by anything
 * counting keys.
 */
export function withCondition(
  conditions: Record<number, string> | undefined,
  slide: number,
  column: string,
): Record<number, string> | undefined {
  const next: Record<number, string> = { ...conditions };
  if (column === "") delete next[slide];
  else next[slide] = column;
  return Object.keys(next).length > 0 ? next : undefined;
}

/**
 * Conditions that name a column the attached data does not have.
 *
 * The engine reports these too (`unknownConditions`) and emits the slide
 * anyway, deliberately, so an authoring mistake is not hidden behind output
 * that looks finished. The pane says it BEFORE the merge as well, because at
 * that point it is still free to fix.
 *
 * Reachable without anyone typing a column name: the control offers only
 * columns that exist, and then the user pastes different data.
 */
export function danglingConditions(state: PaneState): string[] {
  if (!state.columns || !state.conditions) return [];
  const have = new Set(state.columns);
  const out: string[] = [];
  for (const slide of blockSlides(state)) {
    const column = state.conditions[slide];
    if (column !== undefined && !have.has(column) && !out.includes(column)) out.push(column);
  }
  return out;
}

/**
 * The columns whose cells name picture files.
 *
 * Read off the PARSE, not guessed at by the pane: `detectType` decides, the
 * same way it decides a column is a number or a date, so the pane and the
 * engine cannot disagree about which column is an image.
 */
export function imageColumns(state: PaneState): string[] {
  return (state.records?.columns ?? []).filter((c) => c.type === "image").map((c) => c.name);
}

/**
 * The columns this merge will take PICTURES from.
 *
 * Two sources, and the pane only had one. `imageColumns` is what the detector
 * decided, and it is all-or-nothing on purpose: one cell reading `n/a` in a
 * column of file names makes the whole column text, so that a column of `.svg`
 * names is not offered as pictures and then failed one row at a time.
 *
 * The other source is the AUTHOR. `{{Photo|image}}` is documented in
 * `docs/MANUAL.md` as the way to ask for a picture, and the ENGINE obeys it
 * whatever the column's type — it resolves the cell and fills the shape. The
 * pane did not: the picker is shown only when this list is non-empty, so a
 * column one stray cell had kept out of the type gave an author who wrote the
 * format by hand a pane with nowhere to attach their files and a merge that
 * left every picture placeholder standing.
 *
 * `imageFieldsIn` has answered "which fields ask for a picture" since it was
 * written, and until now nothing in the product called it.
 */
export function pictureColumns(state: PaneState): string[] {
  const known = new Set((state.records?.columns ?? []).map((c) => c.name));
  const out = new Set(imageColumns(state));
  // Only a field that names a column the data actually has. One that does not
  // is an unmatched field, which the fields step already reports as such.
  for (const name of state.imageFields ?? []) if (known.has(name)) out.add(name);
  return [...out];
}

/** Every picture file name the attached data refers to, across every picture column. */
export function imagesWanted(state: PaneState): string[] {
  const rows = state.records?.rows ?? [];
  const seen = new Set<string>();
  for (const column of pictureColumns(state)) for (const name of imageNamesIn(rows, column)) seen.add(name);
  return [...seen];
}

export interface ImageTally {
  wanted: number;
  matched: number;
  /** Names the data asks for that no picked file answers to. */
  missing: string[];
  /** Files the user picked that no row refers to. Not a problem — worth saying once. */
  spare: string[];
}

/**
 * What the picked files cover, against what the data asks for.
 *
 * The whole report for the images control, computed in one place so the line
 * under the picker, the merge summary and the warning cannot drift. Matched by
 * BASE NAME and case-insensitively, exactly as `runPlan` matches — a pane that
 * counted matches by a different rule than the merge uses would promise
 * pictures that never arrive.
 */
export function imageTally(state: PaneState): ImageTally {
  const wanted = imagesWanted(state);
  const have = new Set([...(state.images?.keys() ?? [])].map(baseName));
  const missing = wanted.filter((n) => !have.has(baseName(n)));
  const referenced = new Set(wanted.map(baseName));
  const spare = [...(state.images?.keys() ?? [])].filter((f) => !referenced.has(baseName(f)));
  return { wanted: wanted.length, matched: wanted.length - missing.length, missing, spare };
}

/**
 * The `{{…}}` a column is written as on a slide.
 *
 * One function, because this string is produced in three places — the button
 * that inserts it, the text put on the clipboard when the host refuses the
 * insert, and the chip that says which columns are already on the slides — and
 * the engine's own reader (`fieldsIn`) is the fourth party that has to agree.
 * Spelled differently in any one of them and the insert lands a placeholder
 * the merge will never fill.
 *
 * No trimming and no case folding: `fieldsIn` matches the name between the
 * braces exactly, and `unmatchedFields` compares against the column name
 * exactly, so anything done here would have to be done there too.
 */
export function fieldToken(column: string, kind?: string): string {
  // An image column's placeholder asks for a PICTURE, and the engine decides
  // that from the format rather than from the column. Written here so the chip
  // the user presses and the token the engine reads are one string.
  return kind === "image" ? `{{${column}|image}}` : `{{${column}}}`;
}

/**
 * The columns this template can actually carry a field for, and the ones it
 * cannot.
 *
 * `canBeField` is the ENGINE's own reader, asked rather than restated. The
 * first version of the Insert control offered a button per column with no such
 * check, and shipped a defect within the hour: a header holding a brace or a
 * pipe produces a token `FIELD` reads as a different, shorter name, so the
 * button would put a placeholder on the slide that binds to nothing and says
 * nothing.
 *
 * Split rather than filtered, because a column that cannot be a field is worth
 * NAMING — the fix is to rename the column, and a chip that is simply missing
 * tells the user nothing about which one or why.
 */
export function insertableColumns(state: PaneState): { can: string[]; cannot: string[] } {
  const can: string[] = [];
  const cannot: string[] = [];
  for (const column of state.columns ?? []) (canBeField(column) ? can : cannot).push(column);
  return { can, cannot };
}

/**
 * Columns with no placeholder on the slides yet.
 *
 * The mirror of `unmatchedFields`, and the one the fields step is actually
 * about: that names a field with no column behind it, this names a column the
 * template never uses. Neither blocks a merge — a column nobody put on a slide
 * is a perfectly ordinary thing to have in a spreadsheet — so this only ever
 * orders the list and says which chips are already placed.
 */
export function unusedColumns(state: PaneState): string[] {
  if (!state.columns) return [];
  const placed = new Set(state.fields);
  return state.columns.filter((c) => !placed.has(c));
}

/** The chosen block as a sentence subject: "Slide 4", "Slides 4 to 6". */
function blockSubject(state: PaneState): string {
  const block = chosenBlock(state);
  if (!block) return "Those slides";
  return block.from === block.to ? `Slide ${block.from}` : `Slides ${block.from} to ${block.to}`;
}

/**
 * Why there is nothing to merge yet, naming the slides.
 *
 * Kept apart from the engine's own refusal (`prepareBlock`) deliberately: that
 * one has to tell a user with no data and no pane in front of them what to
 * type, so it spells the syntax out. By the time this one is shown the columns
 * are attached and there is a button for each of them, so the advice is to
 * press one — telling somebody to type `{{First}}` by hand next to a button
 * that inserts it is worse advice, not shorter.
 *
 * Shown on PREVIEW and MERGE, where the fix is a step away. The fields step has
 * its own wording, below, because the advice there cannot be "go to the fields
 * step".
 */
export function noFieldsYet(state: PaneState): string {
  return `${blockSubject(state)} carry no fields yet. Go back a step and put one on a slide.`;
}

/**
 * The same fact on the step that fixes it.
 *
 * A separate sentence rather than one with a branch in it, because the two say
 * different things: this one points AT the buttons above it, and the other
 * points at this step. The first version shared `noFieldsYet` and told a user
 * standing on the fields step to go to the fields step — which a screenshot
 * caught and no assertion would have.
 */
export function noFieldsHere(state: PaneState): string {
  return `${blockSubject(state)} carry no fields yet. Press a column above to put one on the slide, then check the slides again.`;
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
    case "data":
      return chosenBlock(state) ? null : "Choose the slides that repeat first.";
    case "fields":
      // Data first, because a field IS a column name. With nothing attached
      // this step has nothing to offer and nothing to check against.
      if (!chosenBlock(state)) return "Choose the slides that repeat first.";
      if (!state.rows) return "Attach your data first.";
      return null;
    case "preview":
      // A live preview is never blocked, because there IS something to do here
      // and the button already does it: remove it. The three sentences below
      // name what is missing before a preview can be MADE, and typing in the
      // slide-number boxes clears the block without ending one — so this step
      // could show "Choose the slides that repeat first." directly above a
      // working "Remove the preview" button, naming the wrong next action at
      // the one moment the user has an obvious right one.
      if (state.previewing) return null;
      if (!chosenBlock(state)) return "Choose the slides that repeat first.";
      if (!state.rows) return "Attach your data first.";
      // A preview runs the ORDINARY merge, which refuses a block with no
      // placeholders — so without this the button would spend a template read
      // and a host insert to arrive at that refusal. Said here instead, where
      // it is still free to fix.
      if (state.fields.length === 0) return noFieldsYet(state);
      return null;
    case "merge": {
      if (!chosenBlock(state)) return "Choose the slides that repeat first.";
      if (!state.rows) return "Attach your data first.";
      // The engine refuses this too, and must: N identical copies is never what
      // anybody meant and is expensive to undo. This is the same rule said
      // before the run rather than after it.
      if (state.fields.length === 0) return noFieldsYet(state);
      // Not the same as having no data. A user who unticked every row has
      // done something deliberate and needs telling what, not "attach data".
      if (includedCount(state) === 0) return "Every row is unticked, so there is nothing to merge.";
      // A placeholder with no column used to be refused here. It is a CAUTION
      // now — see `caution` below for why, and for where the sentence went.
      if (state.previewing) return "End the preview before merging.";
      return null;
    }
  }
}

/**
 * Something true about this run that is worth saying and not worth refusing.
 *
 * Separate from `blockedReason` because the two answer different questions.
 * That one says why a step cannot run; this one says what will happen when it
 * does. Merging them is how a warning becomes a wall.
 *
 * A placeholder with no column was a wall until 2026-08-29, and three other
 * parts of this project already disagreed with it:
 *
 * - the ENGINE leaves such a placeholder on the slide, deliberately, so a half
 *   filled deck does not look finished;
 * - the PREVIEW step ran the ordinary merge with one and produced correct
 *   slides, having no such check of its own;
 * - `docs/MANUAL.md` promises it in as many words — "a row whose picture is
 *   missing keeps its placeholder, exactly as a text field with no column
 *   does".
 *
 * The asymmetry was the tell. A field whose COLUMN is missing was refused; a
 * field whose column exists but whose PICTURE is missing was allowed and
 * documented. Both end with a placeholder on the slide.
 *
 * What the wall was protecting against is real — a typo merged across 240
 * slides is expensive — but the user has already been told twice by here: the
 * fields step outlines the chip and names it in a card, and this sentence sits
 * directly above the button. Being told is the protection; being stopped was
 * not.
 */
export function caution(state: PaneState, step: StepId): string | null {
  if (step !== "merge") return null;
  const missing = unmatchedFields(state);
  // Both can be true at once, and each is a different thing about to happen, so
  // neither may swallow the other.
  const pictures = picturesGoNowhere(state)
    ? "The pictures you attached will not be placed: no field asks for one. Write a field as {{Column|image}} on the slide where the picture goes."
    : null;
  if (missing.length === 0) return pictures;
  // NAMED, not counted — a count alone sends the user back through every slide
  // looking for it. And it says what will HAPPEN rather than what to fix: the
  // placeholder staying is the documented behaviour, not a mistake to correct
  // before pressing on.
  const unmatched = `No column for ${missing.join(", ")}. ${missing.length === 1 ? "It" : "They"} will stay on the slides as written.`;
  return pictures ? `${unmatched} ${pictures}` : unmatched;
}

/**
 * Pictures attached, and nothing on the slides asking for one.
 *
 * The picker is offered from the DATA — a column of file names — and it has to
 * be, because the pane's order is choose slides, paste data, then put the
 * fields on. At paste time the placeholder may not exist yet.
 *
 * The ENGINE decides differently, and correctly: a picture is placed where a
 * field ASKS for one, `{{Photo|image}}`, which is what `docs/MANUAL.md`
 * documents. A plain `{{Photo}}` is a text field whose cell happens to hold a
 * file name, and it merges to `ada.png` on the slide.
 *
 * So a user could attach three files to a picker that said "3 pictures named in
 * Photo", press merge, and get 240 slides reading `ada.png` — with nothing
 * anywhere having mentioned it. The picker is right, the engine is right, and
 * the screen never put the two together.
 *
 * Said only once the slides have been read, because `imageFields` is empty
 * before that and an empty list then means "not looked yet" rather than "not
 * there".
 */
export function picturesGoNowhere(state: PaneState): boolean {
  const attached = (state.images?.size ?? 0) > 0;
  const asked = (state.imageFields ?? []).length > 0;
  return attached && !asked && state.fields.length > 0;
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
  if (state.running === "inspect" && (step === "template" || step === "fields")) {
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
    case "data":
      // "Attach data" is what the step is FOR, so it stays the label until
      // there is data; once there is, the button states what it will carry
      // forward. A button that says "Attach data" after the data is attached
      // reads as a step that did not take.
      return state.rows
        ? { label: `Use ${state.rows} row${state.rows === 1 ? "" : "s"}`, enabled: reachable }
        : { label: "Attach data", enabled: false };
    case "fields":
      // One press, one job: read the slides again and go on. The user has just
      // been putting `{{Column}}` onto them in PowerPoint, and nothing tells
      // this pane that happened — there is no document-changed event for slide
      // text — so the fields it lists are as old as the last read.
      //
      // The label states what is KNOWN rather than what the press does, once
      // anything is known: "Check the slides again" beside a list of six
      // fields reads as a step that has not taken.
      return state.fields.length > 0
        ? {
            label: `Use ${state.fields.length} field${state.fields.length === 1 ? "" : "s"}`,
            enabled: reachable,
          }
        : { label: "Check the slides for fields", enabled: reachable };
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
      const n = slidesToAdd(state);
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
 * Precedence is the states a user must ACT on first — a preview is sitting in
 * their deck and a finished merge is sitting in it too, where an unmatched
 * field is something they can leave alone and a tick is only decoration.
 */
export type OrangeHolder = "tick" | "preview" | "undone" | "unmatched";

export function orangeHolder(state: PaneState, step: StepId): OrangeHolder {
  if (state.previewing) return "preview";
  // A landed merge outranks the tick for the same reason a preview does: the
  // card offering the slides back is the thing on screen worth looking at, and
  // a tick beside it is a second orange saying "done" about the very state the
  // card exists to undo.
  if (step === "merge" && (state.added ?? 0) > 0) return "undone";
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
