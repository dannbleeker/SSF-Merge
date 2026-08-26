/**
 * The pane's DOM, built from the pure decisions and deciding nothing itself.
 *
 * Kept away from Office.js so it can be rendered in the suite: `main.ts` is the
 * only file in this directory allowed to touch the host, and
 * `test/architecture.test.ts` holds that. A renderer that reads
 * `Office.context` mid-build is a screen nobody can check without a PowerPoint,
 * and this is the one surface where a wrong label is what the user acts on.
 *
 * Everything user-visible comes from `steps.ts` and `summary.ts`. Nothing here
 * decides whether a button is enabled or what it says.
 */
import { STEPS, STEP_TITLE, blockedReason, orangeHolder, primary, statusOf, unmatchedFields } from "./steps.js";
import type { OrangeHolder, PaneState, StepId } from "./steps.js";
import { blockSummary, mergeArithmetic, mergeSummary } from "./summary.js";

function el<K extends keyof HTMLElementTagNameMap>(
  doc: Document,
  tag: K,
  props: { class?: string; text?: string; attrs?: Record<string, string> } = {},
): HTMLElementTagNameMap[K] {
  const node = doc.createElement(tag);
  if (props.class) node.className = props.class;
  // textContent, never innerHTML: a column name, a placeholder and a file name
  // all reach this screen from a file somebody pasted.
  if (props.text !== undefined) node.textContent = props.text;
  for (const [k, v] of Object.entries(props.attrs ?? {})) node.setAttribute(k, v);
  return node;
}

/** The four-segment progress rail. */
function rail(doc: Document, state: PaneState, current: StepId): HTMLElement {
  const ul = el(doc, "ul", { class: "rail" });
  for (const step of STEPS) {
    ul.append(
      el(doc, "li", {
        attrs: { "data-status": statusOf(state, step, current), "data-step": step, title: STEP_TITLE[step] },
      }),
    );
  }
  return ul;
}

/**
 * The whole pane for one step.
 *
 * `root` is emptied first: rendering into a half-cleared pane is how a stale
 * count survives a state change and gets pressed.
 */
export function render(root: HTMLElement, state: PaneState, current: StepId): void {
  const doc = root.ownerDocument;
  root.textContent = "";

  const main = el(doc, "main");
  const n = STEPS.indexOf(current) + 1;
  // One orange per view, decided in one place. See orangeHolder.
  const orange = orangeHolder(state, current);

  if (orange === "tick") main.append(el(doc, "span", { class: "tick" }));

  main.append(el(doc, "p", { class: "step-of", text: `Step ${n} of ${STEPS.length} · ${STEP_TITLE[current]}` }));
  main.append(el(doc, "h1", { text: headline(state, current) }));
  main.append(rail(doc, state, current));

  if (state.previewing) {
    const card = el(doc, "div", { class: "card undo" });
    card.append(el(doc, "p", { text: "A preview is on the slide. The template's own text is stored and put back." }));
    main.append(card);
  }

  for (const node of body(doc, state, current, orange)) main.append(node);

  const why = blockedReason(state, current);
  if (why) main.append(el(doc, "p", { class: "blocked", text: why }));

  const action = primary(state, current);
  const button = el(doc, "button", { class: "primary", text: action.label, attrs: { "data-action": current } });
  button.disabled = !action.enabled;
  main.append(button);

  root.append(main);
}

function headline(state: PaneState, current: StepId): string {
  switch (current) {
    case "template":
      return state.block ? blockSummary(state.block, state.rows) : "Which slides repeat?";
    case "fields":
      return state.fields.length === 0 ? "No placeholders found" : `${state.fields.length} placeholders`;
    case "preview":
      return "See one row on the slide";
    case "merge":
      return state.block && state.rows ? mergeArithmetic(state.block, state.rows) : "Nothing to merge yet";
  }
}

function body(doc: Document, state: PaneState, current: StepId, orange: OrangeHolder): HTMLElement[] {
  const out: HTMLElement[] = [];
  if (current === "template") {
    out.push(
      el(doc, "p", {
        class: "muted",
        text: "Pick the first and last slide of the set that should repeat once per row. They must sit next to each other.",
      }),
    );
    return out;
  }

  if (current === "fields") {
    const missing = new Set(unmatchedFields(state));
    const list = el(doc, "ul", { class: "fields" });
    for (const field of state.fields) {
      // The chip is only allowed its orange border when it HOLDS the budget.
      // With a preview showing, the preview card has it and these stay plain.
      const marked = missing.has(field) && orange === "unmatched";
      list.append(el(doc, "li", { text: field, attrs: { "data-matched": marked ? "no" : "yes" } }));
    }
    out.push(list);
    if (state.columns && missing.size > 0) {
      const card = el(doc, "div", { class: "card" });
      card.append(
        el(doc, "p", {
          class: "muted",
          // Named, not counted. A count sends the user back through every slide.
          text: `No column for ${[...missing].join(", ")}.`,
        }),
      );
      out.push(card);
    }
    return out;
  }

  if (current === "merge" && state.block && state.rows) {
    // The heading already states the arithmetic; repeating it here made the
    // screen say "240 rows x 3 slides" twice, which reads as a rendering bug.
    // The card carries the CONSEQUENCE, which is the other half of the answer.
    const card = el(doc, "div", { class: "card summary" });
    card.append(el(doc, "p", { class: "facts", text: mergeSummary(state.block, state.rows, state.deckSize ?? 0) }));
    out.push(card);
  }
  return out;
}
