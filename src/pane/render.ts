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
import {
  EMPTY_DRAFT,
  STEPS,
  STEP_TITLE,
  blockedReason,
  chosenBlock,
  orangeHolder,
  primary,
  readBlockDraft,
  readPastedTable,
  slidesPerRecord,
  statusOf,
  unmatchedFields,
} from "./steps.js";
import type { OrangeHolder, PaneState, StepId } from "./steps.js";
import { blockName, blockSummary, mergeArithmetic, mergeSummary, plural } from "./summary.js";

function el<K extends keyof HTMLElementTagNameMap>(
  doc: Document,
  tag: K,
  props: { class?: string; text?: string; attrs?: Record<string, string>; value?: string } = {},
): HTMLElementTagNameMap[K] {
  const node = doc.createElement(tag);
  if (props.class) node.className = props.class;
  // textContent, never innerHTML: a column name, a placeholder and a file name
  // all reach this screen from a file somebody pasted.
  if (props.text !== undefined) node.textContent = props.text;
  for (const [k, v] of Object.entries(props.attrs ?? {})) node.setAttribute(k, v);
  // The PROPERTY, not the attribute. `setAttribute("value", …)` sets the
  // default an input reverts to, so a re-render mid-typing would put the box
  // back to what it held before the keystroke that caused it.
  if (props.value !== undefined && "value" in node) (node as { value: string }).value = props.value;
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
    // Names the SLIDES, because a user who closes the pane mid-preview has no
    // other way to find out which ones to delete. And says the template is
    // untouched, because the obvious fear about a preview on a real deck is
    // that it has edited the thing being previewed — which is exactly what the
    // design on this project's rejected list would have done.
    card.append(
      el(doc, "p", {
        text: state.previewSlides
          ? `${blockName(state.previewSlides)} ${state.previewSlides.from === state.previewSlides.to ? "is" : "are"} a preview of the first row. Your template is untouched — removing the preview deletes them.`
          : "A preview is in your deck. Your template is untouched — removing the preview deletes it.",
      }),
    );
    main.append(card);
  }

  for (const node of body(doc, state, current, orange)) main.append(node);

  const why = blockedReason(state, current);
  if (why) main.append(el(doc, "p", { class: "blocked", text: why }));

  // What the HOST said, kept separate from what the step needs: "PowerPoint
  // would not name every slide between 4 and 6" is not a thing the user did
  // wrong, and filing it under the same sentence as "Attach your data first"
  // makes both read as nagging.
  if (state.notice) main.append(el(doc, "p", { class: "blocked notice", text: state.notice }));

  // Before the primary, so the primary stays the LAST element in the view —
  // the rule the layout was approved on. Rendered as a link rather than a
  // second button so there is still visibly one thing to press.
  const previous = STEPS[STEPS.indexOf(current) - 1];
  if (previous) {
    main.append(
      el(doc, "button", {
        class: "back",
        text: `Back to ${STEP_TITLE[previous].toLowerCase()}`,
        attrs: { "data-back": previous },
      }),
    );
  }

  // The preview step is the only one whose primary ACTS rather than advances —
  // it shows a row — so it is the only one that needs a way forward that is not
  // the button. Without this the wizard has no exit from step 3 at all, which
  // is what making the preview real took away.
  if (current === "preview" && !state.previewing && blockedReason(state, "merge") === null) {
    main.append(
      el(doc, "button", {
        class: "back forward",
        text: "Skip to the merge",
        attrs: { "data-forward": "merge" },
      }),
    );
  }

  const action = primary(state, current);
  const button = el(doc, "button", { class: "primary", text: action.label, attrs: { "data-action": current } });
  button.disabled = !action.enabled;
  main.append(button);

  root.append(main);
}

function headline(state: PaneState, current: StepId): string {
  switch (current) {
    case "template": {
      const block = chosenBlock(state);
      return block ? blockSummary(block, state.rows) : "Which slides repeat?";
    }
    case "fields":
      // plural(), not a template literal. The zero case was special-cased and
      // the ONE case was not, so a name badge or certificate template holding a
      // single {{Name}} announced "1 placeholders" — and the screenshot script
      // only ever renders three, so nothing showed it either.
      return state.fields.length === 0 ? "No placeholders found" : plural(state.fields.length, "placeholder");
    case "preview":
      return state.previewing ? "The first row is in your deck" : "See one row before you commit";
    case "merge": {
      const block = chosenBlock(state);
      return block && state.rows ? mergeArithmetic(block, state.rows) : "Nothing to merge yet";
    }
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
    out.push(blockControl(doc, state));
    // Reading the selection is safe on this host and it is MEASURED so — see
    // `selectedBlock`. Offered as a link beside the boxes rather than as the
    // primary, because typing two numbers always works and clicking slides is
    // the shortcut: if the host ever stops answering, the step still functions.
    out.push(
      el(doc, "button", {
        class: "back use-selection",
        text: "Use the slides I have selected",
        attrs: { "data-action": "selection" },
      }),
    );
    const read = readBlockDraft(state.draft ?? EMPTY_DRAFT, state.deckSize);
    // A `why` WITH a block is advice the user may press past; without one it is
    // a refusal. Both are shown the same way — the button already says which,
    // by being live or not.
    if (read.why) out.push(el(doc, "p", { class: "blocked", text: read.why }));
    return out;
  }

  if (current === "fields") {
    out.push(dataControl(doc, state));
    const read = readPastedTable(state.paste ?? "");
    if (read.why) out.push(el(doc, "p", { class: "blocked", text: read.why }));
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

  if (current === "preview" && !state.previewing) {
    const previewBlock = chosenBlock(state);
    out.push(
      el(doc, "p", {
        class: "muted",
        text: previewBlock
          ? `Adds ${plural(slidesPerRecord(previewBlock), "slide")} to the end of the deck — the first row, merged the way every row will be. Look at them, then remove them.`
          : "Choose the slides that repeat first.",
      }),
    );
    return out;
  }

  const block = chosenBlock(state);
  // The card is a FORECAST — "720 slides added after slide 12, leaving 732 in
  // the deck" — and it is only true before the run. Left up during the merge it
  // announces the slides as already added while the button reads "Merging…",
  // and left up after it, it says the same sentence the notice says, which is
  // the "does not state the same arithmetic twice" defect one element over.
  // Once a run starts, what actually happened is the notice's to report.
  const forecast = state.running === undefined && state.added === undefined;
  if (current === "merge" && block && state.rows && forecast) {
    // The heading already states the arithmetic; repeating it here made the
    // screen say "240 rows x 3 slides" twice, which reads as a rendering bug.
    // The card carries the CONSEQUENCE, which is the other half of the answer.
    const card = el(doc, "div", { class: "card summary" });
    card.append(el(doc, "p", { class: "facts", text: mergeSummary(block, state.rows, state.deckSize ?? 0) }));
    out.push(card);
  }
  return out;
}

/**
 * The two slide-number boxes.
 *
 * `type="number"` so a phone or a touch keyboard offers digits, and `min` so
 * the spinner cannot walk below the first slide — but the READING is
 * `readBlockDraft`, never the input's own validity. A number input reports ""
 * for a box holding "--", so trusting it would turn nonsense into an empty box
 * and an empty box into a silent zero.
 *
 * `data-field` is what `main.ts` reads. Labels wrap their input rather than
 * using `for`/`id`, because two panes cannot both own an id and this one is
 * re-rendered from scratch on every keystroke.
 */
function blockControl(doc: Document, state: PaneState): HTMLElement {
  const draft = state.draft ?? EMPTY_DRAFT;
  const row = el(doc, "div", { class: "row" });
  for (const [field, caption] of [
    ["from", "First slide"],
    ["to", "Last slide"],
  ] as const) {
    const label = el(doc, "label", { class: "grow field" });
    label.append(el(doc, "span", { class: "caption", text: caption }));
    label.append(
      el(doc, "input", {
        value: draft[field],
        // TEXT, not number, and `inputmode` is what still offers a touch
        // keyboard its digits. `type="number"` refuses `selectionStart` and
        // throws on `setSelectionRange`, so the caret could not be restored
        // across the redraw this pane does on every keystroke — typing 5 into
        // "4|6" gave 456 and the 9 after it went to the end. The reading was
        // never the input's job anyway: `readBlockDraft` is the authority, and
        // it is the only thing that ever saw "0" or "1.5".
        attrs: { type: "text", inputmode: "numeric", autocomplete: "off", "data-field": field },
      }),
    );
    row.append(label);
  }
  return row;
}

/**
 * The paste box.
 *
 * A textarea rather than a file picker because the commonest thing a user has
 * is a range selected in Excel, and Ctrl+C into here is one step where a file
 * picker is four and needs the file saved first. The parse is
 * `readPastedTable`, which sniffs tab first for exactly that input.
 *
 * What it says underneath is the COLUMNS it found, not "data attached": the
 * whole class of failure this step has is a paste that parsed into one column
 * because the copy came through as plain text, and a row count alone looks
 * perfectly healthy when that happens.
 */
function dataControl(doc: Document, state: PaneState): HTMLElement {
  // A LABEL, matching blockControl. As a div the caption was a sibling span
  // that named nothing: a screen reader reached step 2 and announced "edit,
  // multiline, blank", and clicking the caption focused the box on step 1 and
  // did nothing on step 2. `.field` already styles it as a column, so wrapping
  // needs no CSS and introduces no id.
  const wrap = el(doc, "label", { class: "field" });
  wrap.append(el(doc, "span", { class: "caption", text: "Paste your rows, headers included" }));
  wrap.append(
    el(doc, "textarea", {
      value: state.paste ?? "",
      attrs: { rows: "5", spellcheck: "false", "data-field": "paste" },
    }),
  );
  if (state.columns && state.rows) {
    wrap.append(
      el(doc, "p", {
        class: "muted",
        text: `${plural(state.rows, "row")} · ${state.columns.join(", ")}`,
      }),
    );
  }
  return wrap;
}
