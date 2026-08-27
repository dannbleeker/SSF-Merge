/**
 * The pane's entry point, and the only file here that touches Office.js.
 *
 * Everything it shows comes from `render.ts`, everything it decides comes from
 * `steps.ts` and `summary.ts`, and all three are checked by the suite without a
 * PowerPoint anywhere. `test/architecture.test.ts` holds that seam: a decision
 * that migrates into this file becomes untestable the moment it arrives.
 */
import { ready as hostReady, selectedBlock, slideCount } from "../office/powerpoint.js";
import { inspectBlock, runMerge, undoMerge, type MergeOutcome } from "../office/merge.js";
import { render } from "./render.js";
import {
  EMPTY,
  EMPTY_DRAFT,
  chosenBlock,
  firstRowOnly,
  nextStep,
  readPastedTable,
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
function draw(): void {
  const active = document.activeElement;
  const field =
    active instanceof HTMLInputElement || active instanceof HTMLTextAreaElement
      ? active.getAttribute("data-field")
      : null;
  // Read the selection BEFORE the element is destroyed. A number input answers
  // null for both, which is exactly why the boxes are `type="text"`.
  const start = field !== null && active instanceof HTMLElement ? selectionOf(active) : null;

  render(root(), state, step);

  if (field === null) return;
  const node = root().querySelector(`[data-field="${selectorSafe(field)}"]`);
  if (!(node instanceof HTMLInputElement) && !(node instanceof HTMLTextAreaElement)) return;
  node.focus();
  if (start === null) return;
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
    // Only rendered on a step whose primary does not advance, and only when the
    // destination is reachable — `render` asks `blockedReason` before drawing
    // it, so this does not re-decide what the screen already decided.
    step = forward as StepId;
    state = { ...state, notice: undefined };
    draw();
    return;
  }

  const action = target.closest("[data-action]")?.getAttribute("data-action");
  if (!action) return;
  if (action === "merge") {
    void merge();
    return;
  }
  if (action === "template") {
    void useBlock();
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
  if (!(target instanceof HTMLInputElement) && !(target instanceof HTMLTextAreaElement)) return;
  const field = target.getAttribute("data-field");
  if (!field) return;
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
    state = { ...state, draft, block: undefined, fields: [], notice: undefined, added: undefined };
    draw();
    return;
  }

  if (field === "paste") {
    const read = readPastedTable(target.value);
    state = {
      ...state,
      paste: target.value,
      notice: undefined,
      added: undefined,
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
    state = picked.ok
      ? {
          ...state,
          draft: { from: String(picked.from), to: String(picked.to) },
          // `block` cleared so it keeps meaning "a block whose placeholders
          // have been READ". Nothing observable distinguishes this from
          // committing the selection — `chosenBlock` prefers the draft either
          // way, and the template step's only way forward is the button that
          // reads — so it is stated here rather than guarded by a test that
          // would pass against both. Defensive, and known to be.
          block: undefined,
          fields: [],
          added: undefined,
          notice: undefined,
        }
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
async function useBlock(): Promise<void> {
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
      ? { ...state, deckSize, block, fields: report.fields, notice: undefined }
      : { ...state, deckSize, block: undefined, fields: [], notice: report.detail };
    if (report.ok) step = nextStep("template") ?? step;
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

/** A raise as a sentence. */
function readable(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
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
  const records = state.records;
  const conditions = state.conditions;
  if (!block || !state.rows || !records || state.running) return;
  // In the STATE. The first version disabled the button by hand, on a DOM node
  // every later `draw()` replaces with one `primary()` had re-enabled — so a
  // Back and a Continue during a two-minute merge handed the user a live
  // "Add 720 slides" over a run already in flight, and their deck got both.
  state = { ...state, running: "merge", notice: undefined };
  draw();
  try {
    const outcome = await runMerge({
      from: block.from,
      to: block.to,
      records,
      ...(conditions ? { conditions } : {}),
    });
    last = outcome;
    state = {
      ...state,
      deckSize: outcome.deckAtStart + outcome.added,
      // The fields the RUN found, which is the authority: `inspectBlock` read
      // them before the merge and this is the same read after it.
      ...(outcome.fields.length > 0 ? { fields: outcome.fields } : {}),
      // Only a run that ADDED something disarms the button. A refusal that
      // added nothing should leave the user able to press again.
      ...(outcome.added > 0 ? { added: outcome.added } : {}),
      notice: outcome.detail,
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
      last = {
        ok: false,
        detail: readable(e),
        added,
        deckAtStart: before,
        runId: "recovered",
        fields: [],
        unknownConditions: [],
      };
    }
    state = {
      ...state,
      ...(deckAfter !== undefined ? { deckSize: deckAfter } : {}),
      ...(added > 0 ? { added } : {}),
      notice:
        added > 0
          ? `The merge raised, and ${added} slide${added === 1 ? "" : "s"} landed anyway: ${readable(e)}`
          : `The merge did not run: ${readable(e)}`,
    };
  } finally {
    state = { ...state, running: undefined };
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
  const records = state.records;
  if (!block || !records || state.running) return;
  state = { ...state, running: "preview", notice: undefined };
  draw();
  try {
    const outcome = await runMerge({
      from: block.from,
      to: block.to,
      records: firstRowOnly(records),
      ...(state.conditions ? { conditions: state.conditions } : {}),
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
      ...(outcome.fields.length > 0 ? { fields: outcome.fields } : {}),
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

void Office.onReady(() => {
  applyTheme();
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
  draw();
  // The deck's size, so the template boxes can warn about a block that runs
  // past the end before a template read is spent on it. Failing is not fatal:
  // the warning is skipped and `blockIds` still catches it later. It is re-read
  // on every press of "Use slides N to M", because a count taken once at open
  // goes stale the moment the user adds a slide.
  void slideCount().then(
    (deckSize) => {
      state = { ...state, deckSize };
      draw();
    },
    () => undefined,
  );
});

export { applyTheme, advance };
