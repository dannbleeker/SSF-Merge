/**
 * The pane's entry point, and the only file here that touches Office.js.
 *
 * Everything it shows comes from `render.ts`, everything it decides comes from
 * `steps.ts` and `summary.ts`, and all three are checked by the suite without a
 * PowerPoint anywhere. `test/architecture.test.ts` holds that seam: a decision
 * that migrates into this file becomes untestable the moment it arrives.
 */
import { ready as hostReady, slideCount } from "../office/powerpoint.js";
import { inspectBlock, runMerge, type MergeOutcome } from "../office/merge.js";
import { render } from "./render.js";
import { EMPTY, EMPTY_DRAFT, chosenBlock, nextStep, readPastedTable, type PaneState, type StepId } from "./steps.js";

let state: PaneState = EMPTY;
let step: StepId = "template";

function root(): HTMLElement {
  const node = document.getElementById("pane");
  if (!node) throw new Error("the pane's root element is missing");
  return node;
}

function draw(): void {
  render(root(), state, step);
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

  if (field === "from" || field === "to") {
    const draft = { ...(state.draft ?? EMPTY_DRAFT), [field]: target.value };
    // The committed block goes, because it is now stale: leaving it would let
    // `chosenBlock` fall back to slides the boxes no longer name. The fields
    // read off it go with it for the same reason.
    state = { ...state, draft, block: undefined, fields: [], notice: undefined };
    draw();
    // The caret. `render` rebuilds the pane, so the box that was being typed
    // in is a different element by the time this returns.
    focusField(field);
    return;
  }

  if (field === "paste") {
    const read = readPastedTable(target.value);
    state = {
      ...state,
      paste: target.value,
      notice: undefined,
      ...(read.records
        ? { records: read.records, columns: read.columns, rows: read.rows }
        : { records: undefined, columns: undefined, rows: undefined }),
    };
    draw();
    focusField(field);
  }
}

/** Put the caret back where it was, at the end of what is there. */
function focusField(field: string): void {
  const node = root().querySelector(`[data-field="${field}"]`);
  if (!(node instanceof HTMLInputElement) && !(node instanceof HTMLTextAreaElement)) return;
  node.focus();
  const end = node.value.length;
  // A number input throws on setSelectionRange, which is why this is guarded
  // rather than called on both.
  if (node instanceof HTMLTextAreaElement) node.setSelectionRange(end, end);
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
  if (!block) return;
  state = { ...state, notice: "Reading the slides…" };
  draw();
  const report = await inspectBlock({ from: block.from, to: block.to });
  state = report.ok
    ? { ...state, block, fields: report.fields, notice: undefined }
    : { ...state, block: undefined, fields: [], notice: report.detail };
  if (report.ok) advance("template");
  else draw();
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
  if (!block || !state.rows || !state.records) return;
  const button = root().querySelector("button.primary");
  if (button instanceof HTMLButtonElement) {
    button.disabled = true;
    button.textContent = "Merging…";
  }
  const outcome = await runMerge({
    from: block.from,
    to: block.to,
    records: state.records,
    ...(state.conditions ? { conditions: state.conditions } : {}),
  });
  last = outcome;
  // The fields the RUN found, which is the authority: `inspectBlock` read them
  // before the merge and this is the same read after it.
  state = {
    ...state,
    deckSize: outcome.deckAtStart + outcome.added,
    ...(outcome.fields.length > 0 ? { fields: outcome.fields } : {}),
    notice: outcome.detail,
  };
  draw();
}

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
  // The deck's size, so the template boxes can refuse a block that runs past
  // the end before a template read is spent on it. Failing is not fatal: the
  // check is skipped and `prepareBlock` still catches it later.
  void slideCount().then(
    (deckSize) => {
      state = { ...state, deckSize };
      draw();
    },
    () => undefined,
  );
});

export { applyTheme, advance };
