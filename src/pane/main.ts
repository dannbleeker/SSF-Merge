/**
 * The pane's entry point, and the only file here that touches Office.js.
 *
 * Everything it shows comes from `render.ts`, everything it decides comes from
 * `steps.ts` and `summary.ts`, and all three are checked by the suite without a
 * PowerPoint anywhere. `test/architecture.test.ts` holds that seam: a decision
 * that migrates into this file becomes untestable the moment it arrives.
 */
import { ready as hostReady } from "../office/powerpoint.js";
import { runMerge, type MergeOutcome } from "../office/merge.js";
import { render } from "./render.js";
import { EMPTY, type PaneState, type StepId } from "./steps.js";

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
  const action = target.closest("[data-action]")?.getAttribute("data-action");
  if (!action) return;
  if (action === "merge") {
    void merge();
    return;
  }
  advance(action as StepId);
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
  if (!state.block || !state.rows || !state.records) return;
  const button = root().querySelector("button.primary");
  if (button instanceof HTMLButtonElement) {
    button.disabled = true;
    button.textContent = "Merging…";
  }
  const outcome = await runMerge({
    from: state.block.from,
    to: state.block.to,
    records: state.records,
    ...(state.conditions ? { conditions: state.conditions } : {}),
  });
  last = outcome;
  state = { ...state, deckSize: outcome.deckAtStart + outcome.added };
  draw();
  say(outcome.detail);
}

/** The last run, so an undo has the numbers it is clamped against. */
let last: MergeOutcome | undefined;

function say(message: string): void {
  const node = document.createElement("p");
  node.className = "blocked";
  node.textContent = message;
  root().append(node);
}

function advance(from: StepId): void {
  const order: StepId[] = ["template", "fields", "preview", "merge"];
  const next = order[order.indexOf(from) + 1];
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
  root().addEventListener("click", onClick);
  draw();
});

export { applyTheme, advance };
