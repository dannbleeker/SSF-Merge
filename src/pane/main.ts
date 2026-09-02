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
  documentKey,
  hostEnvironment,
  insertTextAtCursor,
  ready as hostReady,
  selectedBlock,
  slideCount,
} from "../office/powerpoint.js";
import { inspectBlock, runMerge, undoMerge, type MergeOutcome } from "../office/merge.js";
import { nextSweepOffer } from "../host/undo.js";
import { readable } from "../host/errors.js";
import { clearCrumb, dropCrumb, readCrumb } from "./crumb.js";
import { blockDrafted, blockMoved, dataChanged } from "./transitions.js";
import type { EmptyPolicy } from "../core/merge/resolve.js";
import { beginRun, onTrace, trace, traceText } from "../core/trace.js";
import { render } from "./render.js";
import { describeMerge, plural } from "./summary.js";
import { upright } from "./upright.js";
import {
  EMPTY,
  EMPTY_DRAFT,
  blockedReason,
  announcement,
  chosenBlock,
  disclosureKey,
  fieldToken,
  pictureColumns,
  firstIncludedRow,
  includedRecords,
  nextStep,
  readPastedTable,
  slidesToAdd,
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
/**
 * How many presses that prove NOTHING the pane will go on offering.
 *
 * One is too few: a host that failed a single tag read, and a host that
 * accepted a delete and performed none, both answer "nothing removed, nothing
 * proved" and both come good on the next press — withdrawing on the first
 * threw away slides that were still removable. Unbounded is too many: a host
 * stuck in that mode leaves a delete button standing over slides no press can
 * take, for the rest of the session.
 *
 * Two is the smallest number that tells a hiccup from a state.
 */
const FRUITLESS_LIMIT = 2;

/**
 * The id of the pending marker THIS pane wrote, if it has written one.
 *
 * The boot deck count can answer minutes after it was issued, by which time the
 * user may have pressed Merge — and the crumb it then reads is that run's own
 * marker. Recognising it by id is what lets the recovery notice stay silent
 * about a run that is happening, without going silent about a genuine one.
 */
let pendingRunId: string | undefined;

/**
 * An id for a crumb written by a run that has no id of its own yet.
 *
 * "pending" is written before the insert answers, "recovered" after a raise —
 * both used to be written as those bare words, and a bare word is not an
 * identity: `dropCrumb` carries a date forward for the same run, and two
 * crashed merges on decks of the same size then shared one, so a merge that
 * failed today was dated weeks ago in the notice offering it back. Anything
 * downstream still treats these as placeholders — `provenSweep` falls back
 * whenever the id is not among the slides' own marks, which no generated id
 * ever is.
 */
function placeholderRunId(kind: "pending" | "recovered"): string {
  return `${kind}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

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

/**
 * The live region, made once and never rebuilt.
 *
 * `render` empties `#pane` and builds fresh elements on every draw, and a live
 * region CREATED with its content in it does not announce — the region has to
 * exist first and have text put into it. So this one lives outside the pane, is
 * made on the first draw, and is only ever written to.
 *
 * Made here rather than in `taskpane.html` so there is one definition and the
 * jsdom wiring tests get it without keeping a copy of the page's markup in step
 * with the real one. It is off-screen rather than `display: none`, which would
 * take it out of the accessibility tree along with everything in it.
 */
function liveRegion(): HTMLElement {
  const existing = document.getElementById("announcer");
  if (existing) return existing;
  const node = document.createElement("p");
  node.id = "announcer";
  node.className = "visually-hidden";
  // `polite`, never `assertive`: none of this is urgent enough to cut across
  // what the user is already being told, and `assertive` on a pane that
  // redraws this often is how a screen reader becomes unusable.
  node.setAttribute("role", "status");
  node.setAttribute("aria-live", "polite");
  document.body.append(node);
  return node;
}

/** The last thing announced, so the same sentence is not said twice. */
let announced = "";

function announce(): void {
  const say = announcement(state);
  // Only on a CHANGE. The pane redraws on every keystroke, and writing the same
  // string back into a live region makes some screen readers say it again.
  if (say === announced) return;
  announced = say;
  liveRegion().textContent = say;
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
  // After the render, so the sentence announced is the one now on screen.
  announce();

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
  // The three collapsible controls, from the one table that knows they exist.
  // Three branches stood here, identical but for the key they flipped, and a
  // fourth control meant remembering to add a fourth.
  const open = disclosureKey(action);
  if (open) {
    state = { ...state, [open]: !state[open] };
    draw();
    return;
  }
  if (action === "selection") {
    void useSelection();
    return;
  }
  // The card's plain way back: take the preview out and STAY on this step.
  if (action === "end-preview") {
    void endPreview();
    return;
  }
  if (action === "preview") {
    // While a preview is up the primary carries ON to the merge, taking the
    // preview out on the way. The merge step refuses while one is showing, so
    // removing it is part of going there rather than a chore to do first.
    void (state.previewing ? endPreviewAndAdvance() : preview());
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
    state = { ...blockDrafted(state, draft), notice: undefined };
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
 * One host job at a time, and the pane says so while it runs.
 *
 * "One host call at a time" is a documented property of this product, not an
 * implementation detail — it is in the manual — and it had a scar behind it
 * before it had a flag: going back a step and forward again during a
 * ninety-second template read handed the user a live "Add 720 slides" over a
 * merge already in flight, and their deck got both. `PaneState.running` is the
 * fix, and until 2026-08-30 it was applied by hand in six places: six guards,
 * six sets on the way in, and six `finally` blocks clearing it.
 *
 * Five of those `finally` blocks were character-for-character identical. That
 * is not a tidiness complaint — the rule that matters is "the button comes back
 * on EVERY path", and a rule written out six times is a rule that holds until
 * somebody writes the seventh without a `finally`.
 *
 * `entering` is whatever else goes into the state in the SAME assignment that
 * raises the flag, and it is one patch rather than a set of named options
 * because the call sites want three different things from `notice` alone: a
 * string SAYS something while the call is out ("Reading the slides…"),
 * `notice: undefined` CLEARS whatever the last run said, and omitting the key
 * keeps what is on screen — which is what ending a preview wants, since the
 * sentence above the button is about the merge and not about the preview being
 * taken back. Spread semantics already express all three, so nothing here has
 * to encode them.
 *
 * Same assignment, and that is the point rather than a convenience: the flag
 * and everything cleared alongside it reach the screen in one draw. The merge
 * clears two more fields on the way in, and doing that in a second statement
 * would put one draw on screen carrying this run's flag beside the last run's
 * waiting line.
 *
 * `whenItRaises` is optional, and its absence means the body handles its own
 * failures. Only the merge does: a raise there does not mean nothing happened,
 * so it has to count the deck and keep the numbers an undo is clamped against,
 * which is not a thing a generic handler can do.
 */
async function duringRun(
  kind: NonNullable<PaneState["running"]>,
  opts: { entering?: Partial<PaneState>; whenItRaises?: (e: unknown) => string },
  body: () => Promise<void>,
): Promise<void> {
  // The busy check lives HERE, so a seventh caller cannot forget it. Callers
  // keep their own preconditions — a block to read, a run to take back —
  // because those are about the caller and this is about the pane.
  if (state.running) return;
  state = { ...state, running: kind, ...opts.entering };
  draw();
  try {
    await body();
  } catch (e) {
    if (!opts.whenItRaises) throw e;
    state = { ...state, notice: opts.whenItRaises(e) };
  } finally {
    state = { ...state, running: undefined };
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
  await duringRun("inspect", { entering: { notice: "Reading your selection…" } }, async () => {
    const picked = await selectedBlock();
    // `block` cleared so it keeps meaning "a block whose placeholders have been
    // READ". Nothing observable distinguishes this from committing the
    // selection — `chosenBlock` prefers the draft either way — so it is stated
    // rather than guarded by a test that would pass against both.
    // Through the same rule the slide-number boxes use. Selecting the SAME
    // slides has not moved the block, so the conditions keyed to it are not
    // stale — and this path calling `blockMoved` directly is how that fix came
    // to cover typing and not selecting.
    state = picked.ok
      ? {
          ...blockDrafted(state, { from: String(picked.from), to: String(picked.to) }),
          notice: undefined,
        }
      : { ...state, notice: picked.why };
  });
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
  if (!block) return;
  await duringRun("inspect", { entering: { notice: "Reading the slides…" }, whenItRaises: readable }, async () => {
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
          imageFieldsOffSlide: report.imageFieldsOffSlide,
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
    //
    // And only if the user is still WHERE THEY ASKED FROM. The read takes
    // seconds, the Back link stays live throughout, and this moved the step
    // whenever the answer arrived: walk back to Data while the template is
    // being read and the pane jumped you forward two steps to Preview, over
    // work you had gone back to change. The block's own staleness is already
    // re-checked above; this is the other half of the same question.
    const next = nextStep(from);
    if (report.ok && next && step === from && blockedReason(state, next) === null) step = next;
  });
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
  // ADDED to what is already held, not put in its place. A browser's picker
  // returns one directory's selection, and a spreadsheet built from a photo
  // library routinely names files in several — so picking the second folder
  // threw the first one away and the tally then reported every name from it as
  // missing, with the files sitting on the disk the author had just chosen
  // them from. There is no other way to attach two folders.
  //
  // Same-name-wins-last inside the map, which is the rule a single pick
  // already follows and the one `clashingPicturesNote` exists to warn about.
  const images = new Map<string, Uint8Array>(state.images ?? []);
  const held = images.size;
  let refused = 0;
  for (const file of Array.from(files)) {
    try {
      // Turned here, once, rather than at merge time. PowerPoint ignores a
      // photo's EXIF orientation — established on a real host on 2026-09-02 —
      // so a phone's portrait picture lands in the deck lying on its side
      // unless the pixels are turned first. `upright` returns the bytes
      // unchanged for everything that does not need it, which is nearly
      // everything, and also whenever it cannot do the work at all.
      images.set(file.name, await upright(new Uint8Array(await file.arrayBuffer())));
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
      : held > 0
        ? // Said only when there was something to add TO, because a reader who
          // expects the picker to replace needs telling once that it did not.
          { notice: `Added to the pictures already attached — ${plural(images.size, "file")} now.` }
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
  if (!block || !records || records.rows.length === 0) return;
  // The flag, and the two fields that belong in the same assignment as it: a
  // draw carrying THIS run's flag beside the LAST run's waiting line is a pane
  // saying it is waiting on a call nobody made. `duringRun` owns the flag —
  // see there for why it is a flag in the state and not a disabled button.
  await duringRun("merge", { entering: { notice: undefined, inFlight: undefined, log: undefined } }, async () => {
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
    /**
     * The deck's size, COUNTED HERE, and the floor every sweep is clamped to.
     *
     * Not `state.deckSize`. That is read when the block is committed, and the
     * step between that and this button is the one where the pane sends the
     * user into PowerPoint to put fields on the slides — so a deck that gained
     * a slide in between leaves the pane's number low, and every clamp in
     * `sweepPlan` compares SIZES. None of them compares freshness. A cache two
     * slides behind turns `{from: 12, count: 3}` into `{from: 10, count: 5}`:
     * the run's three slides and two of the user's own, from a plan that
     * satisfies every guard it passes through.
     *
     * The same shape `crumb.ts` records for the document key, one level out —
     * there the missing question was "which deck", here it is "when".
     */
    let deckBefore: number | undefined;
    try {
      // BEFORE the call that makes an undo necessary. It lives in a module
      // variable, so a tab that dies during the insert leaves the deck holding
      // the slides and the pane unable to take them back. `added` is 0 until
      // the deck answers; the crumb is rewritten with the real number below.
      //
      // No crumb at all when the count fails, rather than one clamped to a
      // number nothing proved. A floor that cannot be shown is a floor that
      // authorises nothing, which is the direction `sweepPlan` already takes.
      deckBefore = await slideCount().catch(() => undefined);
      // NOT over a record of slides that are still in the deck.
      //
      // The pending marker is insurance against a tab that dies mid-insert. An
      // earlier run whose slides nobody has taken back is insurance against the
      // same thing, for slides that are already there — and it is worth more,
      // because those slides exist now and this run's may never. Overwriting it
      // is how a second merge that added nothing left six slides in a deck with
      // no record of them anywhere.
      const outstanding = readCrumb(documentKey());
      const holding = outstanding !== undefined && outstanding.added > 0;
      if (deckBefore !== undefined && !holding) {
        pendingRunId = placeholderRunId("pending");
        dropCrumb({ deckAtStart: deckBefore, added: 0, runId: pendingRunId, doc: documentKey() });
      }
      const outcome = await runMerge({
        from: block.from,
        to: block.to,
        records,
        ...(conditions ? { conditions } : {}),
        ...(state.onEmpty ? { onEmpty: state.onEmpty } : {}),
        // The pictures too, so a preview and a merge produce the same slides.
        ...(state.images ? { images: state.images } : {}),
      });
      // Only a run that ADDED something becomes the one an undo is offered for,
      // and the guard is the same one `state.added` below has always carried.
      //
      // Without it the card and the button read from two different runs: the
      // sentence came from `state.added`, correctly still describing the six
      // slides in the deck, while `last` had been overwritten by a merge that
      // added nothing. The offer stayed on screen, the press swept with the
      // second run's numbers, and it removed nothing — forever. The raise path
      // below already guarded this; the success path did not, and that
      // asymmetry was the whole defect.
      // `accountable` as well as `added`: a run that cannot say which of the
      // new slides are its own must not offer to remove them. The sweep
      // refuses that shape anyway, so the offer was a button that answered
      // "nothing to take back" every time it was pressed.
      if (outcome.added > 0 && outcome.accountable) {
        last = outcome;
        dropCrumb({ deckAtStart: outcome.deckAtStart, added: outcome.added, runId: outcome.runId, doc: documentKey() });
      } else if (!holding) {
        // Clear only what this run itself wrote. A crumb describing slides that
        // are still in the deck is not this run's to throw away.
        clearCrumb(documentKey());
      }
      state = {
        ...state,
        // From where the slides LANDED, not from where the run was planned. The
        // two differ only when something added a slide under the merge, and
        // there this was short by exactly that slide — which then feeds
        // `readBlockDraft`'s "past the end of the deck" advice.
        deckSize: (outcome.landedAfter ?? outcome.deckAtStart) + outcome.added,
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
        ...(outcome.added > 0 && outcome.accountable
          ? {
              added: outcome.added,
              // What the run did, kept apart from what is left to sweep. See
              // `addedByRun`.
              addedByRun: outcome.added,
              deckAtStart: outcome.deckAtStart,
              changedSinceMerge: undefined,
              // A NEW merge, so a withdrawal earned by the last one does not
              // hide this one's way back.
              undoWithdrawn: undefined,
              // This run's slides are not a recovered run's, so the card goes
              // back to living on the merge step. `recovered` says WHERE the
              // offer may be drawn, and it was set at boot and never cleared —
              // so one dead run put a slide-deleting button on every step of
              // the wizard for the rest of the session, the template step
              // included, whose whole job is choosing slides out of the deck
              // that button deletes from.
              recovered: undefined,
            }
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
      // The count taken just before the insert, never the pane's cached one.
      // See `deckBefore`: this number is what a positional delete is clamped
      // against, and a stale one reaches past the run into the user's slides.
      const before = deckBefore;
      // CAPPED at what this merge could possibly have built, exactly as
      // `runMerge` caps the same quantity on the path that returns.
      //
      // The delta alone is not "what this run added" — it is everything that
      // arrived, from anywhere. `sweepPlan` refuses when the deck grew by more
      // than the run added, and that clamp is the one thing keeping a
      // positional delete off a stranger's slides; an uncapped count absorbs
      // the excess, so `grew` and `added` come out equal by construction and
      // the clamp can never fire.
      //
      // This is the LIKELIER path for it, which is why the omission mattered: a
      // co-author's slides landing in a shared deck is the same kind of event
      // that makes a call time out in the first place. Twelve of theirs plus
      // six of ours read as eighteen of ours, and the pane offered to remove
      // all eighteen. Capped, `grew > added` stays true and the sweep declines
      // — which is the answer to give when this run cannot say which slides
      // are its own.
      const couldHaveAdded = slidesToAdd(state);
      const grew = deckAfter !== undefined && before !== undefined ? Math.max(0, deckAfter - before) : 0;
      const added = Math.min(grew, couldHaveAdded);
      if (added > 0 && before !== undefined) {
        const recoveredId = placeholderRunId("recovered");
        dropCrumb({ deckAtStart: before, added, runId: recoveredId, doc: documentKey() });
        last = {
          ok: false,
          detail: readable(e),
          added,
          // The recovery path caps `added` at what the run could have added,
          // so what it offers to sweep is already inside what it can account
          // for. See the cap two lines above.
          accountable: true,
          deckAtStart: before,
          runId: recoveredId,
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
        // `undoWithdrawn` goes with them, exactly as on the success path: a
        // withdrawal is about the merge that earned it, and carrying it into
        // the next one hid the only way back to slides that had just landed.
        ...(added > 0 && before !== undefined
          ? { added, addedByRun: added, deckAtStart: before, changedSinceMerge: undefined, undoWithdrawn: undefined }
          : {}),
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
      //
      // INSIDE the body's own `finally`, which runs before `duringRun`'s: that
      // one clears the flag and draws, so this lands in the same paint rather
      // than needing a draw of its own.
      state = { ...state, inFlight: undefined, log: traceText() };
    }
  });
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
  if (!block || !preview) return;
  await duringRun(
    "preview",
    { entering: { notice: undefined }, whenItRaises: (e) => `The preview did not run: ${readable(e)}` },
    async () => {
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
      //
      // From `landedAfter` — the deck's size when the call actually went out —
      // not from `deckAtStart`, which is its size when the run was PLANNED. A
      // slide arriving in between moves every number here by one, and this card
      // exists so a user who closes the pane can find those slides and delete
      // them by hand: it named the co-author's slide and omitted the last of
      // the preview's own. The merge summary was given this anchor and its
      // sibling one function up was not.
      //
      // `accountable` is deliberately NOT checked here, where the merge step
      // checks it before arming its card. A preview whose run cannot say which
      // slides are its own has still put slides in the deck, and the user has
      // to be told; the button then meets `sweepPlan`'s refusal on its first
      // press and ends the preview with a sentence saying so, which is a better
      // outcome than a preview that is on screen with no way to end it.
      const anchor = outcome.landedAfter ?? outcome.deckAtStart;
      const from = anchor + 1;
      state = {
        ...state,
        previewing: true,
        previewSlides: { from, to: anchor + outcome.added },
        deckSize: anchor + outcome.added,
        ...(outcome.fields.length > 0
          ? { fields: outcome.fields, imageFields: outcome.imageFields, slideFields: outcome.slideFields }
          : {}),
        notice: undefined,
      };
    },
  );
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
  if (!outcome) return;
  await duringRun(
    "undo",
    { entering: { notice: undefined }, whenItRaises: (e) => `The slides could not be removed: ${readable(e)}` },
    async () => {
      const { removed, disowned, detail, unprovable, refusedShape } = await undoMerge(outcome);
      if (removed <= 0) {
        // A refusal is an OUTCOME and the detail already says why — usually that
        // the deck changed underneath the run, which is a thing the user can
        // check and act on rather than a failure of the pane.
        //
        // The deck is RE-COUNTED, and that is what withdraws the offer on a
        // refusal. The card's arithmetic went on believing the size it had at
        // merge time, so a refusal printed "nothing to take back (deck was 12,
        // is 19)" beside a live "Remove slides 13 to 18" that would refuse for
        // ever — two deck sizes on one screen, and a destructive button that
        // could not work. `undoIsPossible` asks `sweepPlan` on every draw, so a
        // truthful count takes the card down by itself, and puts it back if the
        // deck comes back.
        //
        // The press is LATCHED whatever it achieved. `pressed` was set only
        // where slides came out, so a press that moved nothing left the next
        // one looking like a first — no proof asked, the whole positional
        // window taken, and a slide the user made in between deleted. A press
        // that moved nothing is still a press.
        //
        // The card comes down on two answers, and neither is "nothing was
        // removed". A host with no `Slide.tags` at all can never prove
        // anything, so the first such press is the last one worth offering; and
        // a host that could prove and did not gets `FRUITLESS_LIMIT` tries,
        // because a failed tag read and a delete the host swallowed both look
        // exactly like this and both come good on the next press. Withdrawing
        // on the first was a defect of its own: it threw away slides the very
        // next press would have removed.
        //
        // `added` and `deckAtStart` STAY in every case. They say what is in the
        // deck, not what may be pressed: `added` is what disarms the merge
        // button, so clearing it re-armed "Add 6 slides" over six slides that
        // were still there.
        const deckNow = await slideCount().catch(() => undefined);
        // EVERY press that reaches here is fruitless — nothing came out. The
        // first version counted only the presses that also disowned something,
        // which left out the other case `FRUITLESS_LIMIT` was written for: a
        // host that accepts the deletes and performs none answers `removed: 0,
        // disowned: 0`, so the offer stood over slides no press could take for
        // the rest of the session. Resetting belongs on the path that actually
        // removed something, and is done there.
        // A refused SHAPE is not a fruitless press. `sweepPlan` declined before
        // anything was asked of PowerPoint — the deck grew past what this run
        // added — so it says nothing about the host, and spending a press from
        // a budget that exists to tell a hiccup from a state would let a
        // co-author's slide use one up. The card comes down anyway while the
        // deck stays that shape, because `undoIsPossible` asks the same
        // question on every draw.
        const fruitless = refusedShape === true ? (outcome.fruitless ?? 0) : (outcome.fruitless ?? 0) + 1;
        const done = unprovable === true || fruitless >= FRUITLESS_LIMIT;
        last = done ? undefined : { ...outcome, pressed: true, fruitless };
        // The crumb is KEPT — it is the record that stops the next merge
        // overwriting a run whose slides are still in the deck — and marked
        // with what this press learned. Without the mark, every future open of
        // this deck said "the pane closed before you could take them back"
        // about a press that was refused, over a card that dies as soon as it
        // is pressed. See `Crumb.unremovable`.
        dropCrumb({
          deckAtStart: outcome.deckAtStart,
          added: outcome.added,
          runId: outcome.runId,
          doc: documentKey(),
          pressed: true,
          ...(fruitless > 0 ? { fruitless } : {}),
          ...(done ? { unremovable: true } : {}),
        });
        state = {
          ...state,
          notice: done
            ? `Nothing was removed — ${detail}. Delete them from the thumbnail rail if you want them gone.`
            : `Nothing was removed — ${detail}`,
          ...(done ? { undoWithdrawn: true } : {}),
          ...(deckNow !== undefined ? { deckSize: deckNow } : {}),
        };
        return;
      }
      // What a SECOND press may ask for, decided in one place for this screen
      // and the preview's — see `nextSweepOffer`. A press that DECLINED a slide
      // ends the offer, because carrying any count forward re-submits that
      // slide to a window it is now alone in, where `provenSweep` takes an
      // all-untagged plan whole. A press that took slides and declined none may
      // be pressed again.
      const offer = nextSweepOffer({ added: outcome.added, removed, disowned });
      const declined = (disowned ?? 0) > 0;
      // `pressed`, so the next sweep refuses the pre-tags fall-through: the deck
      // has provably changed shape by now, and the window a size clamp produces
      // can hold a slide the user made since. See `provenSweep`.
      // `fruitless: 0`, because this press removed slides. Spreading the old
      // outcome carried the count through a success, so one fruitless press
      // before a working one and one after it withdrew the card on the second
      // — a press short of the budget the changelog and the manual both
      // promise, with merged slides still in the deck.
      last = offer !== null ? { ...outcome, added: offer, pressed: true, fruitless: 0 } : undefined;
      // The slides are the crumb's whole reason. Gone, and it is noise that would
      // offer a stale recovery on the next open.
      // `pressed` travels in the crumb too. The guard it feeds is about the
      // DECK's history rather than this pane session's: close the pane after a
      // partial undo and reopen it, and the next press is still not a first
      // one, because a press has happened and the deck has changed shape.
      if (offer !== null)
        dropCrumb({
          deckAtStart: outcome.deckAtStart,
          added: offer,
          runId: outcome.runId,
          doc: documentKey(),
          pressed: true,
        });
      else clearCrumb(documentKey());
      // ASKED, not computed. The sentence used to say the deck was "back to"
      // the size it started at, which is only true when the sweep took
      // everything — replace two merged slides with two of your own and it said
      // "back to 12" over a deck of 14. Subtracting `removed` from the pane's
      // cached size was the next answer and is wrong in the same place: the
      // cache is stale precisely when the user has edited the deck by hand,
      // which is the only way to reach the branch that prints it. The two
      // sibling paths already re-count; this one guessed.
      const deckNow =
        (await slideCount().catch(() => undefined)) ??
        (state.deckSize ?? outcome.deckAtStart + outcome.added) - removed;
      state = {
        ...state,
        deckSize: deckNow,
        // Only a sweep with nothing left to do disarms the button. One that
        // left slides it CAN still take keeps the way back on screen; one that
        // met a slide it will not claim has to stop offering, because it can no
        // longer tell that slide from its own by position.
        ...(offer !== null
          ? { added: offer, deckAtStart: outcome.deckAtStart }
          : { added: undefined, deckAtStart: undefined }),
        notice:
          offer !== null
            ? `Some of the merge is still there — ${detail}`
            : declined
              ? // The DECLINED slides, counted. "The rest" is every slide left
                // in the range, and on a partial sweep most of them were shown
                // to be the merge's — the press simply stopped when it met one
                // it could not claim. Naming a count says what happened; the
                // sentence before it said something false about slides it had
                // never asked about.
                `${plural(removed, "slide")} removed. ${plural(disowned ?? 0, "slide")} in that range could not ` +
                `be shown to be this merge's and ${(disowned ?? 0) === 1 ? "was" : "were"} left alone — delete ` +
                `them from the thumbnail rail if you want them gone.`
              : `${plural(removed, "slide")} removed. Your deck holds ${deckNow}.`,
      };
    },
  );
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
  if (!outcome) return;
  await duringRun("preview", { whenItRaises: (e) => `The preview could not be removed: ${readable(e)}` }, async () => {
    const { removed, disowned, detail } = await undoMerge(outcome);
    // ONE decision for this screen and the merge undo's, so the two cannot
    // answer differently — which they did, and the difference was a deletion.
    // See `nextSweepOffer`.
    const offer = nextSweepOffer({ added: outcome.added, removed, disowned });
    // How many of the preview this run still believes are in the deck. It
    // decides what to SAY; `offer` decides whether pressing again is worth
    // offering, and the two are different questions — a slide the sweep
    // declined is still in the deck and is not this run's to take.
    const outstanding = Math.max(0, outcome.added - removed - (disowned ?? 0));
    if (outstanding > 0) {
      // ASK THE DECK before claiming anything. The commonest way to reach here
      // is that the user did what the card told them the button does — deleted
      // the preview slides themselves — so the deck never grew, the sweep
      // refused, and this said "Some of the preview is still there" about
      // slides that are gone.
      //
      // And every branch out of here CLEARS `previewing` unless another press
      // could finish the job. While a preview is up the forward link is
      // withheld, the rail is not clickable and the merge step refuses, so a
      // preview that cannot be ended is a pane with no way on — a state this
      // screen has now reached by three separate routes.
      const deckNow = await slideCount().catch(() => undefined);
      if (offer !== null) {
        // Slides came out, none was declined, and some are still owed, so
        // pressing again is worth offering. That is NOT the same as the window
        // holding only this run's slides — this comment said it was, which is
        // the claim `nextSweepOffer`'s docstring exists to retract: the user may
        // have deleted merged slides by hand and made others, and every quantity
        // here is a size. The next press asks `provenSweep` for proof, which is
        // what makes the offer safe rather than the window being clean.
        //
        // The card stops NAMING them. After a partial removal it cannot: the
        // ones that went took the numbering of the ones that stayed with them,
        // so it went on saying "Slides 5 to 8 are a preview of the first row"
        // over a deck where 5 is the user's own slide again, beside a button
        // offering to delete them.
        shown = { ...outcome, added: offer, pressed: true };
        state = {
          ...state,
          previewSlides: undefined,
          notice: `Some of the preview is still there — ${detail}`,
          ...(deckNow !== undefined ? { deckSize: deckNow } : {}),
        };
        return;
      }
      // Nothing more this pane may do. Either the sweep declined a slide — and
      // carrying any count forward re-submits it to a window it is now alone
      // in, where an all-untagged plan is taken whole — or the press moved
      // nothing and would answer the same way again.
      //
      // The sentence may not say the preview is gone. `deckNow` is a SIZE:
      // a deck back to where it started is equally the user having deleted
      // their own slides, with preview slides still in it.
      shown = undefined;
      state = {
        ...state,
        previewing: false,
        previewSlides: undefined,
        ...(deckNow !== undefined ? { deckSize: deckNow } : {}),
        notice:
          (removed > 0
            ? // The slides the sweep DECLINED, counted. "The rest" spoke for
              // every slide left in the range, most of which it never doubted
              // — the same sentence the merge undo carried and the same reason
              // it was replaced there.
              `${plural(removed, "slide")} of the preview removed. ${plural(disowned ?? 0, "slide")} could not ` +
              `be shown to be this run's. `
            : `The preview could not be taken back — ${detail}. `) +
          // The size is REPORTED, never read as an identity. A sentence saying
          // the deck is no bigger than it was before the preview was here for
          // one commit and it is the mistake the comment twelve lines above
          // forbids: delete three of your own slides while a three-slide
          // preview is up and the deck is back to its old size with every
          // preview slide still in it — and that branch suppressed the only
          // advice that would have helped. Both, always: the count, and what
          // to do if the slides are still there.
          (deckNow !== undefined ? `Your deck holds ${plural(deckNow, "slide")}. ` : "") +
          "Any preview slides still in your deck can be deleted from the thumbnail rail.",
      };
      return;
    }
    shown = undefined;
    state = {
      ...state,
      previewing: false,
      previewSlides: undefined,
      // What the sweep DISOWNED is still in the deck — it is simply not the
      // preview's, so the deck does not come all the way back. Counting it is
      // arithmetic on numbers already in hand rather than another host call,
      // and a deck size one too small makes the next undo's clamp refuse a
      // sweep it should allow.
      deckSize: outcome.deckAtStart + (disowned ?? 0),
      // Silent when the whole preview came out, which is the ordinary case and
      // wants no sentence. NOT silent when the sweep declined a slide: it is
      // still in the deck, in the range the preview was on, and the user is
      // about to look at a deck with a slide in it they did not expect and no
      // account of where it came from.
      // COULD NOT BE SHOWN, not "are not this run's". `disowned` cannot tell
      // the two apart — a slide with no tag and a slide the host would not
      // answer for are the same `undefined` — and its own docstring says
      // nothing downstream may read it as ownership. On a host below
      // PowerPointApi 1.3 the stronger sentence is false for every user, every
      // time, over slides that ARE this run's.
      notice:
        (disowned ?? 0) > 0
          ? `${removed > 0 ? `${plural(removed, "slide")} of the preview removed. ` : "None of the preview came back. "}` +
            `${plural(disowned ?? 0, "slide")} in that range could not be shown to be this run's and ` +
            `${(disowned ?? 0) === 1 ? "was" : "were"} left alone.`
          : undefined,
    };
  });
}

/**
 * Take the preview out, and go on to the merge if it actually went.
 *
 * Guarded on the state rather than on the call returning, because `endPreview`
 * swallows a partial sweep into a notice: a removal that left some of the
 * preview behind leaves `previewing` true and says so. Advancing anyway would
 * put the user on a merge step refusing with "End the preview before merging."
 * while the sentence explaining what actually happened sat one screen back.
 */
async function endPreviewAndAdvance(): Promise<void> {
  await endPreview();
  if (!state.previewing) advance("preview");
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
      const crumb = readCrumb(documentKey());
      if (crumb && crumb.added > 0 && crumb.unremovable === true) {
        // A press already established that this add-in cannot take these slides
        // back — see `Crumb.unremovable`. Offering the button again on every
        // open, under "the pane closed before you could take them back", is a
        // false sentence over a card that dies the moment it is pressed. The
        // slides ARE there, so the user is told where they are and what to do,
        // and the record is kept so the next merge does not overwrite it.
        state = {
          ...state,
          notice:
            `A merge from ${crumb.startedAt.slice(0, 10)} added ${plural(crumb.added, "slide")} that this add-in ` +
            `could not take back. They are at the end of your deck — delete them from the thumbnail rail if you ` +
            `want them gone.`,
        };
      } else if (crumb && crumb.added > 0) {
        last = {
          ok: false,
          detail: "recovered from a run that did not finish",
          added: crumb.added,
          // The crumb records what a finished run could account for; the sweep
          // re-checks it against the deck and the run tag at press time.
          accountable: true,
          // Carried through, so a run already swept once does not get the
          // pre-tags fall-through on the press after a reopen.
          ...(crumb.pressed === true ? { pressed: true } : {}),
          // The budget too, or a reopen hands a stuck host two more presses.
          ...(crumb.fruitless !== undefined ? { fruitless: crumb.fruitless } : {}),
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
          // `plural`, not the host layer's `slide(s)`. This sentence is on a
          // user's screen; that spelling is a house convention for the run log.
          //
          // "The pane closed before you could take them back" is only true of a
          // crumb NOBODY pressed. `unremovable` was already carved out above,
          // but that mark needs the whole budget spent or a host with no tags
          // at all; `pressed` is written on every fruitless press and is far
          // commoner. The real-host round of 2026-09-02 walked into it: six
          // slides merged, all six deleted from the thumbnail rail, one press
          // answered "nothing to take back (deck was 13, is 13)" — and the next
          // open said the pane had closed before the user could press, beside
          // no card at all, because `sweepPlan` had correctly withdrawn it.
          //
          // A pressed crumb is still worth keeping and still worth mentioning:
          // it is what stops the next merge overwriting the record. It just
          // must not claim the press never happened, or that slides are waiting
          // when the card beside it says otherwise.
          notice:
            crumb.pressed === true
              ? `A merge from ${crumb.startedAt.slice(0, 10)} added ${plural(crumb.added, "slide")}, and a take-back has already been tried on this deck.`
              : `A merge from ${crumb.startedAt.slice(0, 10)} added ${plural(crumb.added, "slide")} and the pane closed before you could take them back.`,
        };
      } else if (crumb && crumb.runId !== pendingRunId) {
        // A run that died DURING the insert, which is the window the crumb was
        // built for and the one it served least: it is written with `added: 0`
        // before the call, and the tab never comes back to write the real
        // number.
        //
        // NOT the marker THIS pane's own run has just written. The deck count
        // this handler waits on is issued at boot and can answer minutes later —
        // long enough for the user to have walked the wizard and pressed Merge —
        // and the crumb it then found was that run's pending marker. It told the
        // user the merge "did not finish" while it was visibly running, and
        // deleted the record in the one window it exists for.
        //
        // Named rather than gated on `state.running`, which was the first fix
        // and is too wide: ANY run in flight — the "Reading the slides…" a user
        // is most likely to start first — suppressed a genuine crash record for
        // the whole session, because this branch is never retried.
        // `placeholderRunId` makes the marker's id unique, so the pane can
        // recognise its own.
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
        clearCrumb(documentKey());
      }
      draw();
    },
    () => undefined,
  );
});

export { applyTheme, advance };
