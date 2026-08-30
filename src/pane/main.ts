/**
 * The pane's entry point, and the only file here that touches Office.js.
 *
 * Everything it shows comes from `render.ts`, everything it decides comes from
 * `steps.ts` and `summary.ts`, and all three are checked by the suite without a
 * PowerPoint anywhere. `test/architecture.test.ts` holds that seam: a decision
 * that migrates into this file becomes untestable the moment it arrives.
 */
import {
  canReadSelection,
  hostEnvironment,
  insertTextAtCursor,
  ready as hostReady,
  selectedBlock,
  slideCount,
} from "../office/powerpoint.js";
import { inspectBlock, runMerge, undoMerge, type MergeOutcome } from "../office/merge.js";
import { readable } from "../host/errors.js";
import { clearCrumb, dropCrumb, readCrumb } from "./crumb.js";
import { blockMoved, dataChanged } from "./transitions.js";
import type { EmptyPolicy } from "../core/merge/resolve.js";
import { beginRun, onTrace, trace, traceText } from "../core/trace.js";
import { render } from "./render.js";
import { describeMerge, plural } from "./summary.js";
import {
  EMPTY,
  EMPTY_DRAFT,
  blockedReason,
  chosenBlock,
  fieldToken,
  pictureColumns,
  firstIncludedRow,
  includedRecords,
  nextStep,
  readPastedTable,
  withCondition,
  type PaneState,
  type StepId,
} from "./steps.js";

let state: PaneState = EMPTY;
let step: StepId = "template";

function root(): HTMLElement {
  const node = document.getElementById("pane");
  if (!node) throw new Error("the pane's root element is missing");
  return node;
}

/**
 * Redraw, keeping the caret where the user left it.
 *
 * `render` empties the root and builds fresh elements, so every draw destroys
 * whatever was focused. That is invisible when a draw follows a click and
 * ruinous when it follows a keystroke: the pane redraws on every character, so
 * without this the caret jumps to the end of the box after each one and
 * "4|6" typed into becomes 4569 instead of 4596. The paste box is worse, because
 * an edit in the middle of a pasted table scatters the rest of the line to the
 * end — and `readPastedTable` then merges the corrupted text.
 *
 * It applies to draws the user did NOT cause, too. The deck count resolves a
 * second or two after the pane opens and redraws; before this, that redraw
 * blanked the focus and swallowed the next digit typed, leaving the box holding
 * what looked like a dropped keystroke.
 */
/**
 * The attributes that say a control is the SAME control after a redraw.
 *
 * Deliberately not `data-back` or `data-forward`: those change the step, so
 * the element with that attribute on the screen afterwards is a different link
 * pointing somewhere else, and focusing it would be a guess rather than a
 * restore.
 *
 * `data-action` covers the primary, the two disclosures, the undo button and
 * the selection shortcut. The primary's value is the STEP, so it restores when
 * the press did not advance — "Preview the first row" becoming "Remove the
 * preview" — and matches nothing when it did, which is the honest answer.
 */
const FOCUS_KEYS = ["data-field", "data-condition", "data-empty", "data-row", "data-insert", "data-action"] as const;

/** A selector for the element focused now, or null if it has no stable name. */
function focusedSelector(active: Element | null): string | null {
  if (!(active instanceof HTMLElement)) return null;
  for (const key of FOCUS_KEYS) {
    const value = active.getAttribute(key);
    if (value !== null) return `[${key}="${selectorSafe(value)}"]`;
  }
  return null;
}

function draw(): void {
  const active = document.activeElement;
  const selector = focusedSelector(active);
  // Read the selection BEFORE the element is destroyed. A number input answers
  // null for both, which is exactly why the boxes are `type="text"`, and a
  // checkbox raises rather than answering — `selectionOf` catches that.
  const start =
    active instanceof HTMLInputElement || active instanceof HTMLTextAreaElement ? selectionOf(active) : null;

  render(root(), state, step);

  if (selector === null) return;
  const node = root().querySelector(selector);
  if (!(node instanceof HTMLElement)) return;
  node.focus();
  if (start === null) return;
  if (!(node instanceof HTMLInputElement) && !(node instanceof HTMLTextAreaElement)) return;
  try {
    node.setSelectionRange(start[0], start[1]);
  } catch {
    // Some input types refuse the selection API outright. Focused is still
    // better than not, and losing the caret position is not worth a raise that
    // would take the whole redraw with it.
  }
}

/** The caret, or null when this element will not say. */
function selectionOf(node: HTMLElement): [number, number] | null {
  const el = node as HTMLInputElement;
  try {
    return el.selectionStart === null || el.selectionEnd === null ? null : [el.selectionStart, el.selectionEnd];
  } catch {
    return null;
  }
}

/**
 * A field name, safe inside a QUOTED attribute selector.
 *
 * `field` comes off a `data-field` attribute this file wrote, so it is one of
 * three known strings today — but it is read back out of the DOM, and a
 * selector built by interpolation is one renamed control away from being a
 * parse error that takes the redraw with it.
 *
 * Deliberately not `CSS.escape`: that escapes an IDENTIFIER, which is not what
 * sits between the quotes here, and the first version of this function reached
 * it as a detached reference — `const escape = CSS?.escape` — which throws
 * `'escape' called on an object that is not a valid instance of CSS` the moment
 * it is called. That took the whole focus restore with it, so the caret went
 * back to being lost on every keystroke, silently, with the tests for it
 * failing on the symptom rather than the cause.
 */
function selectorSafe(value: string): string {
  return value.replace(/["\\]/g, "\\$&");
}

/**
 * Follow PowerPoint's theme, not the browser's.
 *
 * The pane lives inside PowerPoint, which can be dark while the OS is light, so
 * `prefers-color-scheme` is the wrong question. `officeTheme` answers the right
 * one. Outside a host it is undefined — which is the case every time this pane
 * is opened in a browser to look at it — and the stylesheet's media query
 * carries that fallback.
 *
 * Read ONCE, on ready. There is no theme-change event to subscribe to: the
 * typings put `OfficeThemeChanged` on Outlook's `Mailbox` and nowhere else, so
 * a PowerPoint pane cannot be told the theme moved. Switching PowerPoint's
 * theme mid-session therefore needs the pane reopened, which is a real
 * limitation and is written down in the manual rather than hidden behind a
 * handler that never fires. (The first version of this file registered
 * `Office.addin.onVisibilityModeChanged` as though it were a theme event. It is
 * not — it is about the pane being shown and hidden.)
 */
function applyTheme(): void {
  const body = Office.context?.officeTheme?.bodyBackgroundColor;
  if (!body) return;
  const hex = body.replace("#", "");
  const n = Number.parseInt(hex.length === 3 ? [...hex].map((c) => c + c).join("") : hex, 16);
  if (Number.isNaN(n)) return;
  const luminance = ((n >> 16) & 255) * 0.299 + ((n >> 8) & 255) * 0.587 + (n & 255) * 0.114;
  document.documentElement.setAttribute("data-theme", luminance < 128 ? "dark" : "light");
}

function onClick(event: Event): void {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  const back = target.closest("[data-back]")?.getAttribute("data-back");
  if (back) {
    // Straight there, with no reachability check: a step the user has already
    // been through is one they may go back and change, and that is the whole
    // point of the link. Forward is what `nextStep` gates.
    step = back as StepId;
    state = { ...state, notice: undefined };
    draw();
    return;
  }

  const forward = target.closest("[data-forward]")?.getAttribute("data-forward");
  if (forward) {
    // Only rendered on a step whose primary does not advance. Deliberately NOT
    // conditional on the destination being reachable: gating it that way is
    // what left step 4 with no exit when a placeholder had no column. A step
    // that cannot run yet still draws its own `blockedReason` and its own way
    // back, so walking onto it is how the user is TOLD, not how they get stuck.
    step = forward as StepId;
    state = { ...state, notice: undefined };
    draw();
    return;
  }

  // A row checkbox. Read before `data-action`, because the box sits inside the
  // picker and a `closest` for an action would walk past it to the toggle.
  const row = target.closest("[data-row]")?.getAttribute("data-row");
  if (row !== null && row !== undefined) {
    toggleRow(Number(row));
    return;
  }

  // An Insert chip. Read before `data-action` for the same reason the row
  // checkbox is: it sits inside the fields step, whose primary carries one.
  const insert = target.closest("[data-insert]")?.getAttribute("data-insert");
  if (insert !== null && insert !== undefined) {
    void insertField(insert);
    return;
  }

  const action = target.closest("[data-action]")?.getAttribute("data-action");
  if (!action) return;
  if (action === "merge") {
    void merge();
    return;
  }
  if (action === "template" || action === "fields") {
    void useBlock(action);
    return;
  }
  if (action === "rows") {
    state = { ...state, rowsOpen: !state.rowsOpen };
    draw();
    return;
  }
  if (action === "conditions") {
    state = { ...state, conditionsOpen: !state.conditionsOpen };
    draw();
    return;
  }
  if (action === "empties") {
    state = { ...state, emptiesOpen: !state.emptiesOpen };
    draw();
    return;
  }
  if (action === "selection") {
    void useSelection();
    return;
  }
  if (action === "preview") {
    void (state.previewing ? endPreview() : preview());
    return;
  }
  if (action === "undo") {
    void undoRun();
    return;
  }
  advance(action as StepId);
}

/**
 * What the boxes hold, on every keystroke.
 *
 * The draft is stored as TYPED and read by `readBlockDraft`, so a box the user
 * is halfway through is a state the pane can hold rather than a number it has
 * to guess at. The data box is parsed the same way, on input, so the columns
 * appear as the paste lands: `readPastedTable` is the only parse, and what the
 * merge runs on is what those labels were counted from.
 */
function onInput(event: Event): void {
  const target = event.target;
  // A SELECT as well as a box. `<select>` fires `input` as well as `change` in
  // every browser this ships to, so one listener covers both — but the type
  // guard is what decides, and it excluded selects until the condition control
  // needed one.
  if (target instanceof HTMLSelectElement) {
    // What a blank cell does. Its own branch rather than a shared one: it
    // carries no slide number and changes a single field, where a condition is
    // keyed by the slide it is about.
    if (target.hasAttribute("data-empty")) {
      if (state.running) return;
      state = {
        ...state,
        onEmpty: target.value as EmptyPolicy,
        // A different answer here is a different merge — under "skip" it is a
        // different number of slides — so the finished run's disarmed button
        // goes with it, the rule every other edit on this screen follows.
        changedSinceMerge: true,
        notice: undefined,
      };
      draw();
      return;
    }
    const slide = target.getAttribute("data-condition");
    if (slide === null || state.running) return;
    state = {
      ...state,
      conditions: withCondition(state.conditions, Number(slide), target.value),
      // A changed condition changes what the merge produces, so the finished
      // run's disarmed button goes with it — the same rule every other edit on
      // this screen follows. `added` stays: those slides are in the deck, and
      // the undo card is the only thing offering them back.
      changedSinceMerge: true,
      notice: undefined,
    };
    draw();
    return;
  }
  if (!(target instanceof HTMLInputElement) && !(target instanceof HTMLTextAreaElement)) return;
  const field = target.getAttribute("data-field");
  if (!field) return;
  // The picture picker. Handled before the running guard on purpose: reading
  // files touches no host and cannot go stale against an answer in flight, and
  // refusing it would leave the browser's own file dialog having visibly taken
  // a choice the pane then ignored.
  if (field === "images" && target instanceof HTMLInputElement) {
    void takeImages(target.files);
    return;
  }
  // A host call is out and its answer is about to be written into this state.
  // Taking the keystroke would make the answer stale before it lands, which is
  // the race `useBlock`'s staleness check exists to notice.
  if (state.running) return;

  if (field === "from" || field === "to") {
    const draft = { ...(state.draft ?? EMPTY_DRAFT), [field]: target.value };
    // The committed block goes, because it is now stale: leaving it would let
    // `chosenBlock` fall back to slides the boxes no longer name. The fields
    // read off it go with it for the same reason, and `added` goes because a
    // changed block is a different merge.
    state = { ...blockMoved(state), draft, notice: undefined };
    draw();
    return;
  }

  if (field === "rowSearch") {
    state = { ...state, rowSearch: target.value };
    draw();
    return;
  }

  if (field === "paste") {
    const read = readPastedTable(target.value);
    state = {
      ...dataChanged(state),
      paste: target.value,
      notice: undefined,
      ...(read.records
        ? { records: read.records, columns: read.columns, rows: read.rows }
        : { records: undefined, columns: undefined, rows: undefined }),
    };
    draw();
  }
}

/**
 * Fill the two boxes from the slides the user has selected.
 *
 * Fills the DRAFT rather than committing a block, so the numbers land in the
 * boxes and the user still presses "Use slides N to M" — the read that finds
 * the placeholders. Two steps for one action would be worse; silently skipping
 * the template read would be worse still, because the fields step would then
 * show nothing.
 */
async function useSelection(): Promise<void> {
  if (state.running) return;
  state = { ...state, running: "inspect", notice: "Reading your selection…" };
  draw();
  try {
    const picked = await selectedBlock();
    // `block` cleared so it keeps meaning "a block whose placeholders have been
    // READ". Nothing observable distinguishes this from committing the
    // selection — `chosenBlock` prefers the draft either way — so it is stated
    // rather than guarded by a test that would pass against both.
    state = picked.ok
      ? { ...blockMoved(state), draft: { from: String(picked.from), to: String(picked.to) }, notice: undefined }
      : { ...state, notice: picked.why };
  } finally {
    state = { ...state, running: undefined };
    draw();
  }
}

/**
 * Commit the block, and find out what is actually in it.
 *
 * One template read per press, not per keystroke. `inspectBlock` does the same
 * read and the same preparation the merge does and stops before the plan, so
 * the placeholders the fields step lists are the ones the merge will bind —
 * not a guess, and not a second parser that can disagree with the first.
 */
async function useBlock(from: StepId): Promise<void> {
  const block = chosenBlock(state);
  if (!block || state.running) return;
  state = { ...state, running: "inspect", notice: "Reading the slides…" };
  draw();
  try {
    // The deck's size, re-read while we are spending a round trip anyway. It
    // was taken once when the pane opened and never again, so a user who added
    // slides and came back was being told about a deck that no longer existed.
    const deckSize = await slideCount().catch(() => state.deckSize);
    const report = await inspectBlock({ from: block.from, to: block.to });

    // What the boxes say NOW. The read takes seconds and the pane stays on
    // screen; committing a block the user has since retyped puts one block's
    // placeholders behind another block's slides, and `chosenBlock` prefers
    // the draft — so the merge would run on slides nobody read.
    const still = chosenBlock({ ...state, deckSize });
    if (!still || still.from !== block.from || still.to !== block.to) {
      state = { ...state, deckSize, notice: "The slides changed while that was reading. Press again." };
      return;
    }

    state = report.ok
      ? {
          ...state,
          deckSize,
          block,
          fields: report.fields,
          imageFields: report.imageFields,
          // Per slide, because `skippedRows` cannot answer from the flat list.
          slideFields: report.slideFields,
          notice: undefined,
          // The note said "press Check the slides", and this is that press.
          // The list of fields below is the answer now.
          fieldNote: undefined,
        }
      : { ...blockMoved(state), deckSize, notice: report.detail };
    // From whichever step asked. The template step reads to COMMIT a block;
    // the fields step reads again because the user has just been typing
    // `{{Column}}` into PowerPoint and nothing tells this pane that happened —
    // there is no document-changed event for slide text. Same call, same
    // staleness check, and each goes on to its own next step.
    //
    // Only to a step the read has actually UNBLOCKED. A template with no
    // `{{fields}}` on it is a fine answer now — `inspectBlock` no longer
    // refuses one — so a bare `nextStep` would walk a user from the fields
    // step onto a preview that cannot run, where the only thing to do is come
    // back. `blockedReason` is the same question the screen asks, so the pane
    // cannot advance onto a step it would immediately draw as blocked.
    const next = nextStep(from);
    if (report.ok && next && blockedReason(state, next) === null) step = next;
  } catch (e) {
    // `inspectBlock` answers rather than raising, so a raise here is something
    // below it. Said out loud, because the alternative is a pane that reads
    // "Reading the slides…" for the rest of the session.
    state = { ...state, notice: readable(e) };
  } finally {
    // In a `finally`, so the button comes back on every path. A flag cleared
    // only on the happy path is a pane that has to be reopened.
    state = { ...state, running: undefined };
    draw();
  }
}

/**
 * Read the pictures the user picked.
 *
 * Bytes, held in the pane and handed to the engine as a map. Nothing is
 * uploaded and nothing is fetched: the files come off the user's own disk
 * through the browser's picker, which is the only route a sandboxed
 * cross-origin iframe has to them and — for a merge whose whole premise is that
 * the data does not leave — the right one anyway.
 *
 * Keyed by the file's own name. `runPlan` matches a cell against it by base
 * name and case-insensitively, so `Photos\\ada.PNG` in a spreadsheet finds
 * `ada.png` from a folder.
 */
async function takeImages(files: FileList | null): Promise<void> {
  if (!files || files.length === 0) return;
  const images = new Map<string, Uint8Array>();
  let refused = 0;
  for (const file of Array.from(files)) {
    try {
      images.set(file.name, new Uint8Array(await file.arrayBuffer()));
    } catch {
      // A file the browser will not read — moved, or on a disconnected drive.
      // Counted rather than thrown: the others are still worth having, and the
      // tally will name what is still missing anyway.
      refused++;
    }
  }
  state = {
    ...state,
    images,
    // A different set of pictures is a different merge, so a landed run's
    // disarmed button goes with it — the rule every other edit here follows.
    changedSinceMerge: true,
    ...(refused > 0
      ? { notice: `${refused} of the ${files.length} file(s) could not be read and were left out.` }
      : { notice: undefined }),
  };
  draw();
}

/**
 * One press of an Insert chip, and the two ways it can land.
 *
 * `setSelectedDataAsync` puts the token where the cursor is, which is the
 * whole feature — the user clicks into a text box on the slide and presses a
 * column. It is a Common API with no requirement set, so there is nothing to
 * declare and nothing to check: whether this host will do it is found out by
 * asking, once, per press.
 *
 * The FALLBACK is the reason this is worth building at all. On a host that
 * refuses, or with no insertion point, the token goes on the clipboard instead
 * and the pane says so — one Ctrl+V away from the same result, and still
 * better than asking the user to spell a column name from memory with the data
 * in another window. Where even the clipboard is refused (a task pane is a
 * cross-origin iframe and `navigator.clipboard` is gated on permissions there)
 * the sentence carries the token itself, so it can be read off the screen. All
 * three outcomes name the token; none of them is silence.
 *
 * A module flag rather than `state.running`, deliberately. `running` disables
 * the whole pane and relabels its primary, which is right for a two-minute
 * merge and wrong for a chip the user presses six times in a row — the screen
 * would flicker through "Reading the slides…" on every one. This only stops the
 * same press arriving twice.
 */
let inserting = false;

async function insertField(column: string): Promise<void> {
  // A host call is out. Its answer is about to be written into this state, and
  // an insert would be reported against a screen that is about to change.
  if (state.running || inserting) return;
  inserting = true;
  // The IMAGE form for an image column, so the chip and the token agree.
  const token = fieldToken(column, pictureColumns(state).includes(column) ? "image" : undefined);
  try {
    const done = await insertTextAtCursor(token);
    if (done.ok) {
      state = {
        ...state,
        // Says what to do NEXT, because the insert lands on the slide and not
        // in the pane: without the second sentence the user has no reason to
        // press the primary, and the fields list stays empty until they do.
        fieldNote: `${token} put on the slide. Press "Check the slides" when you have placed them all.`,
      };
      return;
    }
    const copied = await copyText(token);
    state = {
      ...state,
      fieldNote: copied
        ? `PowerPoint would not type it in, so ${token} is on your clipboard — click into a text box on the slide and paste it. (${done.why})`
        : `PowerPoint would not type it in, and the clipboard was refused too. Type ${token} onto the slide by hand. (${done.why})`,
    };
  } finally {
    inserting = false;
    draw();
  }
}

/**
 * The clipboard, or false.
 *
 * Never raises. A task pane is a nested cross-origin iframe, where
 * `navigator.clipboard` is gated on a permission the host may not have granted
 * and is absent outright on older WebViews — and this is the FALLBACK path, so
 * a rejection here must produce the next sentence rather than take the handler
 * with it.
 */
async function copyText(text: string): Promise<boolean> {
  try {
    if (!navigator.clipboard?.writeText) return false;
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    return false;
  }
}

/**
 * Take a row out of the merge, or put it back.
 *
 * Stored as EXCLUDED indices, so an untouched state means "merge everything"
 * — which is what a user who never opens the list wants. Refused while a host
 * call is out, like every other input: the answer on its way back is about the
 * rows as they were when it left.
 */
function toggleRow(index: number): void {
  if (!Number.isInteger(index)) return;
  if (state.running) {
    // REDRAW, not a bare return. A checkbox flips itself before the handler
    // runs — that is the control's own default action, in jsdom and in every
    // browser — so refusing without redrawing leaves the box visually unticked
    // while the state says the row is still in. The next draw would put it
    // back, and there is no next draw until the host answers.
    draw();
    return;
  }
  const out = new Set(state.excluded ?? []);
  if (out.has(index)) out.delete(index);
  else out.add(index);
  state = { ...state, excluded: [...out].sort((a, b) => a - b), changedSinceMerge: true };
  draw();
}

/**
 * The merge, and the one thing this file adds to it: telling the user.
 *
 * Every decision is `runMerge`'s. The button is disabled unless the step is
 * reachable, so this does not re-check what `steps.ts` already answered — two
 * copies of that rule is how they come apart.
 *
 * `outcome.deckAtStart` and `outcome.added` are kept because undo is positional
 * and clamped against them. A run whose numbers are lost cannot be taken back
 * safely, so they are held before anything is shown.
 */
async function merge(): Promise<void> {
  const block = chosenBlock(state);
  // The rows the user left ticked, never everything they pasted.
  const records = includedRecords(state);
  const conditions = state.conditions;
  if (!block || !records || records.rows.length === 0 || state.running) return;
  // In the STATE. The first version disabled the button by hand, on a DOM node
  // every later `draw()` replaces with one `primary()` had re-enabled — so a
  // Back and a Continue during a two-minute merge handed the user a live
  // "Add 720 slides" over a run already in flight, and their deck got both.
  state = { ...state, running: "merge", notice: undefined, inFlight: undefined, log: undefined };
  // A run of its own, so the record is this merge and not this session. Pairing
  // one run's numbers with another run's failures is the wrong turn that costs
  // an hour.
  beginRun();
  // AFTER the mark, never at wiring time. A sibling project's environment line
  // was emitted when the pane loaded and the run's slice began later, so it
  // reached NONE of its archived rounds — present in the code, absent from
  // every artefact anyone read.
  trace("pane", "run starting", { ...hostEnvironment(), deck: state.deckSize ?? "unknown" });
  // The window and the record, both kept current as the run goes.
  //
  // The record used to be written ONCE, in the `finally` below, and this
  // comment used to explain that "a run log can only be handed over once the
  // run ends, and the runs worth explaining are the ones that never do". The
  // first half was never true: the entries are in memory from the first call,
  // and `traceText` will format them at any moment. Only the second half was —
  // and it is the argument for writing the record continuously, not for
  // withholding it.
  //
  // What it cost: on a host that wedges, the pane sat on "Waiting on
  // PowerPoint: inserting the merged deck…" forever with nothing to copy, on
  // exactly the run somebody needs to explain. A task pane has no devtools and
  // cannot hand over a file, so what is on screen is the only channel there is.
  //
  // Cheap enough to do on every entry: a merge emits about ten, and the cap is
  // 500. This is not a hot loop.
  onTrace((e) => {
    const call = e.scope === "host" && e.message === "issued" ? e.data?.call : undefined;
    state = { ...state, log: traceText(), ...(typeof call === "string" ? { inFlight: call } : {}) };
    draw();
  });
  // Seeded, not left to the first entry that arrives after the line above.
  // `onTrace` only sees what is written AFTER it subscribes, and the run's
  // first line — the environment, which is the most useful thing in the record
  // — is written before it. A run whose first host call is slow would then
  // show an empty record for as long as that call takes, which is the window
  // this whole change is about.
  state = { ...state, log: traceText() };
  draw();
  try {
    // BEFORE the call that makes an undo necessary. `deckAtStart` is the floor
    // every sweep is clamped to and it lives in a module variable, so a tab
    // that dies during the insert leaves the deck holding the slides and the
    // pane unable to take them back. `added` is 0 until the deck answers; the
    // crumb is rewritten with the real number below.
    dropCrumb({ deckAtStart: state.deckSize ?? 0, added: 0, runId: "pending" });
    const outcome = await runMerge({
      from: block.from,
      to: block.to,
      records,
      ...(conditions ? { conditions } : {}),
      ...(state.onEmpty ? { onEmpty: state.onEmpty } : {}),
      // The pictures too, so a preview and a merge produce the same slides.
      ...(state.images ? { images: state.images } : {}),
    });
    last = outcome;
    if (outcome.added > 0) dropCrumb({ deckAtStart: outcome.deckAtStart, added: outcome.added, runId: outcome.runId });
    else clearCrumb();
    state = {
      ...state,
      deckSize: outcome.deckAtStart + outcome.added,
      // The fields the RUN found, which is the authority: `inspectBlock` read
      // them before the merge and this is the same read after it.
      ...(outcome.fields.length > 0
        ? { fields: outcome.fields, imageFields: outcome.imageFields, slideFields: outcome.slideFields }
        : {}),
      // Only a run that ADDED something disarms the button. A refusal that
      // added nothing should leave the user able to press again.
      // `deckAtStart` travels with `added` everywhere, because the undo card
      // asks `sweepPlan` and a positional offer needs both: `added` says how
      // many slides, this says which.
      ...(outcome.added > 0
        ? { added: outcome.added, deckAtStart: outcome.deckAtStart, changedSinceMerge: undefined }
        : {}),
      // What the merge DID. `outcome.detail` says how much the deck GREW,
      // which is equally true of a merge that filled every placeholder and one
      // that matched none of them and inserted 720 copies of the template.
      notice: outcome.ok ? describeMerge(outcome) : outcome.detail,
    };
  } catch (e) {
    // A raise does not mean nothing happened. This host takes calls it does
    // not perform AND performs calls it then raises on — `insertDeck` reads
    // the deck DELTA for exactly that reason, and its own re-count sits
    // outside the try it uses to do so, so a timeout there rejects with the
    // slides already in the deck. Count again and keep the numbers, because an
    // undo is positional and clamped against them: a run whose numbers are
    // lost cannot be taken back at all.
    const deckAfter = await slideCount().catch(() => undefined);
    const before = state.deckSize;
    const added = deckAfter !== undefined && before !== undefined ? Math.max(0, deckAfter - before) : 0;
    if (added > 0 && before !== undefined) {
      dropCrumb({ deckAtStart: before, added, runId: "recovered" });
      last = {
        ok: false,
        detail: readable(e),
        added,
        deckAtStart: before,
        runId: "recovered",
        fields: [],
        imageFields: [],
        slideFields: [],
        unknownConditions: [],
      };
    }
    state = {
      ...state,
      ...(deckAfter !== undefined ? { deckSize: deckAfter } : {}),
      // `deckAtStart` travels with `added`, here as everywhere: the undo card
      // is a positional offer and needs both. It did not, so the one branch
      // written for a host that performs a call and then raises on it was the
      // one branch that left the slides in the deck with no way to remove
      // them.
      ...(added > 0 && before !== undefined ? { added, deckAtStart: before, changedSinceMerge: undefined } : {}),
      notice:
        added > 0
          ? `The merge raised, and ${added} slide${added === 1 ? "" : "s"} landed anyway: ${readable(e)}`
          : `The merge did not run: ${readable(e)}`,
    };
  } finally {
    // Stop watching before the last draw: the window is for the wait, and a
    // subscriber left attached would relabel `inFlight` on the next run's
    // first call while this run's summary is still on screen.
    onTrace(undefined);
    // Banked whatever happened, and ESPECIALLY when it did not work — a task
    // pane has no devtools a user can open, so the only way this record
    // reaches anybody is by being on screen to copy.
    state = { ...state, running: undefined, inFlight: undefined, log: traceText() };
    draw();
  }
}

/**
 * Show one row on the real slides, and take it back.
 *
 * The preview is the ORDINARY merge over a one-row set, inserted by the same
 * call and removed by the same positional sweep an undo uses. Nothing is
 * written to the template, which is the whole point: setting a shape's text
 * through Office.js re-authors it — office-js#5858, custom bullets reverting
 * to default discs — and the template is the one slide this product exists to
 * preserve.
 *
 * The design this replaced is in the backlog and in this repo's own REJECTED
 * list at the same time: "write one record's values onto the real template
 * slide, store what was there in a tag, and put it back". Putting it back goes
 * through the same API that did the damage, so the text would return and the
 * formatting would not — silently, to the master copy every merged slide is
 * cloned from.
 *
 * What the user sees is therefore produced by the code that will produce the
 * other 239 slides. A preview that renders by some other route is a preview of
 * something nobody is going to get.
 */
async function preview(): Promise<void> {
  const block = chosenBlock(state);
  // The first row that will actually MERGE, not the first that was pasted.
  const preview = firstIncludedRow(state);
  if (!block || !preview || state.running) return;
  state = { ...state, running: "preview", notice: undefined };
  draw();
  try {
    const outcome = await runMerge({
      from: block.from,
      to: block.to,
      records: preview,
      ...(state.conditions ? { conditions: state.conditions } : {}),
      // The blank-cell answer too. A preview run under a different policy from
      // the merge is a preview of something nobody is going to get, which is
      // the rule `firstIncludedRow` already follows for an unticked row.
      ...(state.onEmpty ? { onEmpty: state.onEmpty } : {}),
      // A preview is the ORDINARY merge over one row, so it gets the pictures
      // too. Without them the preview shows a placeholder where the real merge
      // shows a photo, which is a preview of something nobody is going to get.
      ...(state.images ? { images: state.images } : {}),
    });
    if (!outcome.ok || outcome.added === 0) {
      state = { ...state, notice: outcome.detail };
      return;
    }
    shown = outcome;
    // Where they landed, so the card can name them. The insert is anchored
    // after the last slide, so they are the last `added` in the deck.
    const from = outcome.deckAtStart + 1;
    state = {
      ...state,
      previewing: true,
      previewSlides: { from, to: outcome.deckAtStart + outcome.added },
      deckSize: outcome.deckAtStart + outcome.added,
      ...(outcome.fields.length > 0
        ? { fields: outcome.fields, imageFields: outcome.imageFields, slideFields: outcome.slideFields }
        : {}),
      notice: undefined,
    };
  } catch (e) {
    state = { ...state, notice: `The preview did not run: ${readable(e)}` };
  } finally {
    state = { ...state, running: undefined };
    draw();
  }
}

/**
 * Take back the slides the last merge added.
 *
 * The same call and the same clamps as removing a preview, and deliberately so
 * — a preview IS a one-row merge here, so there is one sweep with one set of
 * guarantees rather than two that can drift apart.
 *
 * What differs is WHEN it is pressed. A preview is removed seconds after it
 * lands; this is a button a user may press after looking through 720 slides,
 * on a deck AutoSave has been writing to and a co-author may have been
 * editing. `sweepPlan` refuses outright when the deck has gained more than the
 * run added, because at that point the last N slides are somebody else's and
 * the run's own are unreachable by position — so the worst case here is a
 * refusal with a sentence, never a wrong deletion.
 */
async function undoRun(): Promise<void> {
  const outcome = last;
  if (!outcome || state.running) return;
  state = { ...state, running: "undo", notice: undefined };
  draw();
  try {
    const { removed, detail } = await undoMerge(outcome);
    if (removed <= 0) {
      // A refusal is an OUTCOME and the detail already says why — usually that
      // the deck changed underneath the run, which is a thing the user can
      // check and act on rather than a failure of the pane.
      state = { ...state, notice: `Nothing was removed — ${detail}` };
      return;
    }
    const remaining = outcome.added - removed;
    last = remaining > 0 ? { ...outcome, added: remaining } : undefined;
    // The slides are the crumb's whole reason. Gone, and it is noise that would
    // offer a stale recovery on the next open.
    if (remaining > 0) dropCrumb({ deckAtStart: outcome.deckAtStart, added: remaining, runId: outcome.runId });
    else clearCrumb();
    state = {
      ...state,
      deckSize: (state.deckSize ?? outcome.deckAtStart + outcome.added) - removed,
      // Only a COMPLETE sweep disarms the button. A partial one leaves slides
      // in the deck and the user is the only one who can finish the job, so
      // the way back has to stay on screen.
      ...(remaining > 0
        ? { added: remaining, deckAtStart: outcome.deckAtStart }
        : { added: undefined, deckAtStart: undefined }),
      notice:
        remaining > 0
          ? `Some of the merge is still there — ${detail}`
          : `${plural(removed, "slide")} removed. Your deck is back to ${outcome.deckAtStart}.`,
    };
  } catch (e) {
    state = { ...state, notice: `The slides could not be removed: ${readable(e)}` };
  } finally {
    state = { ...state, running: undefined };
    draw();
  }
}

/**
 * Take the preview back.
 *
 * `undoMerge` is positional and clamped against the count taken before the
 * preview was inserted, so it cannot reach a slide the user owned first — the
 * same guarantee, and the same code, as undoing a real merge.
 */
async function endPreview(): Promise<void> {
  const outcome = shown;
  if (!outcome || state.running) return;
  state = { ...state, running: "preview" };
  draw();
  try {
    const { removed, detail } = await undoMerge(outcome);
    if (removed < outcome.added) {
      // Said out loud rather than assumed. A sweep that removed fewer slides
      // than it asked for leaves some of the preview in the deck, and the user
      // is the only one who can finish the job.
      state = { ...state, notice: `Some of the preview is still there — ${detail}` };
      return;
    }
    shown = undefined;
    state = {
      ...state,
      previewing: false,
      previewSlides: undefined,
      deckSize: outcome.deckAtStart,
      notice: undefined,
    };
  } catch (e) {
    state = { ...state, notice: `The preview could not be removed: ${readable(e)}` };
  } finally {
    state = { ...state, running: undefined };
    draw();
  }
}

/** The preview currently on the slides, so it can be taken back exactly. */
let shown: MergeOutcome | undefined;

/** The last run, so an undo has the numbers it is clamped against. */
let last: MergeOutcome | undefined;

function advance(from: StepId): void {
  const next = nextStep(from);
  if (next) step = next;
  draw();
}

export { last as lastRun };

/**
 * The build this pane was served from, in the header, before anything is run.
 *
 * It already reaches the RUN RECORD through `hostEnvironment()` — but a run
 * record exists only once a run has finished, and the question this answers is
 * asked before one starts: PowerPoint caches the pane's HTML for about ten
 * minutes, so opening it too soon after a deploy tests code the host never
 * fetched, and the result reads as a clean run of the wrong build.
 *
 * That is not hypothetical here. The build immediately before this one never
 * loaded Office.js and showed a header and nothing else; anyone testing the fix
 * has the broken one cached, and the two are told apart by this line or not at
 * all.
 *
 * In the HEADER rather than in the pane, because the layout rule the pane was
 * approved on is that the primary button is the last element in the view. A
 * footer would take that away for a diagnostic.
 */
function showBuild(): void {
  const build = hostEnvironment().build;
  const header = document.querySelector("header");
  if (!header || !build || build === "unknown") return;
  const span = document.createElement("span");
  span.className = "build";
  span.textContent = build;
  // Named, because six hex characters in a header is a mystery otherwise.
  span.title = `SSF Merge was built from commit ${build}`;
  header.append(span);
}

void Office.onReady(() => {
  applyTheme();
  // Before the floor check: a host that cannot run the add-in is exactly the
  // case where somebody needs to say which build refused them.
  showBuild();
  const check = hostReady();
  if (!check.ok) {
    // Said out loud rather than swallowed: a pane that renders a dead UI on an
    // unsupported host gives the user nothing to report.
    const node = root();
    node.textContent = "";
    const p = document.createElement("p");
    p.className = "blocked";
    p.textContent = check.detail;
    node.append(p);
    return;
  }
  const node = root();
  node.addEventListener("click", onClick);
  node.addEventListener("input", onInput);
  // Asked once, before the first draw: it is a version check, not a host call.
  state = { ...state, canSelect: canReadSelection() };
  draw();
  // The deck's size, so the template boxes can warn about a block that runs
  // past the end before a template read is spent on it. Failing is not fatal:
  // the warning is skipped and `blockIds` still catches it later. It is re-read
  // on every press of "Use slides N to M", because a count taken once at open
  // goes stale the moment the user adds a slide.
  void slideCount().then(
    (deckSize) => {
      state = { ...state, deckSize };
      // A run that never finished, offered back.
      //
      // The crumb only matters once the deck's real size is in hand: the offer
      // is a positional sweep, and `sweepPlan` refuses outright if the deck has
      // moved on since. So this is read HERE rather than at boot.
      //
      // This comment claimed the refusal was already the sweep's own, and for a
      // while it was not: the card asked a formula of its own that counted back
      // from the end of the deck, and on a deck that had moved it named the
      // wrong slides — sometimes slides that pre-date the merge — under a
      // button that then did nothing. `undoSummary` and `undoIsPossible` call
      // `sweepPlan` now, which is why `deckAtStart` is carried into the state
      // beside `added`: a positional offer needs both.
      const crumb = readCrumb();
      if (crumb && crumb.added > 0) {
        last = {
          ok: false,
          detail: "recovered from a run that did not finish",
          added: crumb.added,
          deckAtStart: crumb.deckAtStart,
          runId: crumb.runId,
          fields: [],
          imageFields: [],
          slideFields: [],
          unknownConditions: [],
        };
        state = {
          ...state,
          added: crumb.added,
          deckAtStart: crumb.deckAtStart,
          // The offer follows the slides rather than the step, because this
          // pane is on step 1 and the run it is about is over. Without it the
          // sentence below named a button nothing rendered.
          recovered: true,
          notice: `A merge from ${crumb.startedAt.slice(0, 10)} added ${crumb.added} slide(s) and the pane closed before you could take them back.`,
        };
      } else if (crumb) {
        // A run that died DURING the insert, which is the window the crumb was
        // built for and the one it served least: it is written with `added: 0`
        // before the call, and the tab never comes back to write the real
        // number.
        //
        // Told, not offered. The slides may well be in the deck, but nothing
        // here knows how many: taking the deck's growth as the answer would
        // sweep whatever has been appended since, which is exactly the
        // inference `sweepPlan` refuses. So the user gets the two numbers and
        // the place to look, and does it themselves.
        const grew = (deckSize ?? crumb.deckAtStart) - crumb.deckAtStart;
        state = {
          ...state,
          notice:
            `A merge from ${crumb.startedAt.slice(0, 10)} did not finish — the pane closed while PowerPoint was ` +
            `taking the slides. Your deck had ${crumb.deckAtStart} slide(s) before it and has ${deckSize ?? "an unknown number of"} now` +
            `${grew > 0 ? `, so ${grew} may be from that run` : ""}. Check the end of the deck.`,
        };
        // Said once. There is no action attached to it, so leaving the crumb
        // would repeat the same sentence on every open for ever.
        clearCrumb();
      }
      draw();
    },
    () => undefined,
  );
});

export { applyTheme, advance };
