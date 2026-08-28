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
  blockSlides,
  blockedReason,
  chosenBlock,
  conditionFor,
  danglingConditions,
  fieldToken,
  imageColumns,
  imageTally,
  imagesWanted,
  includedCount,
  insertableColumns,
  noFieldsHere,
  unusedColumns,
  rowIncluded,
  rowLabel,
  orangeHolder,
  primary,
  readBlockDraft,
  readPastedTable,
  slidesPerRecord,
  statusOf,
  visibleRows,
  unmatchedFields,
} from "./steps.js";
import type { OrangeHolder, PaneState, StepId } from "./steps.js";
import {
  blockName,
  blockSummary,
  mergeArithmetic,
  mergeSummary,
  plural,
  undoIsPossible,
  undoSummary,
} from "./summary.js";

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

/** The progress rail, one segment per step. */
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

  // What a finished merge added, and the way back.
  //
  // `undoInsert` and `sweepPlan` were built and tested before this and were
  // reachable from nothing: `state.added` was set, `undoSummary` was written
  // and covered, and no view rendered either. So the first real merge into
  // somebody's deck had no way back short of pressing Ctrl+Z the right number
  // of times, and a mail merge is exactly the operation you want to undo after
  // looking at what it produced.
  //
  // Phrased as the SLIDES rather than as "undo", because this deletes part of
  // a presentation and the sentence should say which part.
  if (
    state.added &&
    state.deckSize !== undefined &&
    undoIsPossible(state.added, state.deckSize) &&
    current === "merge"
  ) {
    const card = el(doc, "div", { class: "card undo" });
    card.append(el(doc, "p", { text: undoSummary(state.added, state.deckSize) }));
    card.append(
      el(doc, "button", {
        class: "secondary",
        text: state.running === "undo" ? "Removing…" : "Remove these slides",
        attrs: { "data-action": "undo", ...(state.running ? { disabled: "" } : {}) },
      }),
    );
    main.append(card);
  }

  // WHICH host call is in flight, while one is.
  //
  // A merge is legitimately silent for up to two and a half minutes —
  // `BUDGET.file` allows ninety seconds to read the template and
  // `BUDGET.insert` sixty to hand the package over — and a frozen "Merging…"
  // for that long is indistinguishable from a pane that has wedged. It is also
  // the one thing a user can report that tells us WHERE it stopped.
  if (state.running && state.inFlight) {
    main.append(el(doc, "p", { class: "inflight", text: `Waiting on PowerPoint: ${state.inFlight}…` }));
  }

  // The run record, and the only way it reaches anybody.
  //
  // A task pane is a nested cross-origin iframe with no devtools a user can
  // open, and blob downloads from one are blocked in WebView2
  // (office-js#1511). So the record is on screen, selectable, between markers
  // — the same channel the host probe already uses, because it is the one that
  // works.
  //
  // Shown DURING a run as well as after it. It was gated on the run being over
  // — "so it never competes with the sentence saying what happened" — and
  // there is no such sentence while a run is out: the notice is cleared and
  // the only line on screen is the call being waited on. What the gate did
  // cost is the case the record exists for. A host that wedges never reaches
  // the `finally` that used to write the log, so the run nobody can explain
  // was the run with nothing to copy.
  if (state.log) {
    const details = el(doc, "details", { class: "runlog" });
    details.append(
      el(doc, "summary", {
        // Says which it is. "did" on a run still going is a small lie that
        // makes a reader trust the last line as the final one.
        text: state.running ? "What this run has done so far" : "What this run did, step by step",
      }),
    );
    details.append(el(doc, "pre", { text: `=== SSF MERGE RUN LOG ===\n${state.log}\n=== END ===` }));
    main.append(details);
  }

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
  // the button. Without this the wizard has no exit from step 4 at all, which
  // is what making the preview real took away.
  //
  // Step 1 used to need one too, and no longer does. The template read refused
  // a block with no `{{fields}}` on it, so a fresh deck had no route forward at
  // all and a link was bolted on saying "attach data first to see your column
  // names" — an instruction to go BACKWARDS through a wizard. The order is the
  // fix: data comes before fields now, and the template step advances on the
  // two slide numbers alone.
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
    case "data":
      return state.rows ? plural(state.rows, "row") + " attached" : "Paste the rows to merge";
    case "fields":
      // plural(), not a template literal. The zero case was special-cased and
      // the ONE case was not, so a name badge or certificate template holding a
      // single {{Name}} announced "1 placeholders" — and the screenshot script
      // only ever renders three, so nothing showed it either.
      //
      // The zero case is no longer a dead end: this step now hands the user a
      // button per column, so the heading names the job rather than the
      // absence.
      return state.fields.length === 0 ? "Put your fields on the slides" : plural(state.fields.length, "placeholder");
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
    // Only where the host HAS it. `getSelectedSlides` is PowerPointApi 1.5 and
    // the floor is 1.2, so this is an extra: on an older host the two boxes
    // still work and the shortcut is simply absent, which is better than a
    // button that always fails.
    //
    // Reading the selection is also SAFE where it exists, and that is measured
    // rather than assumed — see `selectedBlock`. Offered as a link rather than
    // as the primary, because typing two numbers always works.
    if (state.canSelect) {
      out.push(
        el(doc, "button", {
          class: "back use-selection",
          text: "Use the slides I have selected",
          attrs: { "data-action": "selection" },
        }),
      );
    }
    const read = readBlockDraft(state.draft ?? EMPTY_DRAFT, state.deckSize);
    // A `why` WITH a block is advice the user may press past; without one it is
    // a refusal. Both are shown the same way — the button already says which,
    // by being live or not.
    if (read.why) out.push(el(doc, "p", { class: "blocked", text: read.why }));
    return out;
  }

  if (current === "data") {
    out.push(dataControl(doc, state));
    const read = readPastedTable(state.paste ?? "");
    if (read.why) out.push(el(doc, "p", { class: "blocked", text: read.why }));
    // The pictures the data asks for. On the DATA step and not the fields step
    // because they are data: the cell names a file, and everything the merge
    // consumes is collected in one place.
    if (imagesWanted(state).length > 0) out.push(imageControl(doc, state));
    return out;
  }

  if (current === "fields") {
    out.push(insertControl(doc, state));
    // What the last Insert press did. Kept out of `notice` on purpose: a notice
    // is cleared by the next edit, and this has to survive the user leaving the
    // pane, clicking into a text box on the slide, and coming back. It is also
    // the only report the clipboard fallback has — an insert lands visibly on
    // the slide, a copy lands nowhere the user can see.
    if (state.fieldNote) out.push(el(doc, "p", { class: "muted note", text: state.fieldNote }));
    const missing = new Set(unmatchedFields(state));
    const list = el(doc, "ul", { class: "fields" });
    for (const field of state.fields) {
      // The chip is only allowed its orange border when it HOLDS the budget.
      // With a preview showing, the preview card has it and these stay plain.
      const marked = missing.has(field) && orange === "unmatched";
      list.append(el(doc, "li", { text: field, attrs: { "data-matched": marked ? "no" : "yes" } }));
    }
    if (state.fields.length > 0) {
      out.push(el(doc, "p", { class: "muted", text: "On the slides now:" }));
      out.push(list);
    } else if (!state.fieldNote) {
      // Not silence. The step's own primary re-reads the slides, and a user who
      // has just inserted three fields and sees nothing has no way to tell an
      // insert that did not land from a pane that has not looked yet.
      //
      // Withheld while a note is up, because between the two the note is the
      // one that is CURRENT. This sentence is read off the last template read,
      // and an insert lands on the slide without telling the pane — so a
      // freshly inserted field puts "{{City}} put on the slide" directly above
      // "these slides carry no fields yet", which is the screen contradicting
      // itself about something the user just did. The note already ends by
      // asking for the read that settles it.
      out.push(el(doc, "p", { class: "muted", text: noFieldsHere(state) }));
    }
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
    card.append(el(doc, "p", { class: "facts", text: mergeSummary(block, includedCount(state), state.deckSize ?? 0) }));
    out.push(card);
  }

  if (current === "merge") {
    // Which rows, and which slides. The two questions the merge screen is for,
    // both shut by default and both summarised shut.
    //
    // The conditions used to live on the fields step, which had the columns
    // they name. That step is now about putting placeholders onto slides, and
    // a control deciding which slides come out for which rows is a different
    // job on the same screen. Here it sits beside the row picker, where the
    // pair reads as "what comes out of this merge".
    //
    // Each keeps its OWN precondition rather than sharing one. A row list needs
    // the rows; a condition needs only the column names, and folding it into
    // the row picker's guard made it vanish from a state that has columns and
    // no records — which is exactly the state a screenshot fixture holds, and
    // is how this was noticed.
    if (state.records && state.rows) out.push(rowPicker(doc, state));
    if (state.columns && state.columns.length > 0) out.push(conditionPicker(doc, state));
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

/**
 * The pictures a merge needs, and where they come from.
 *
 * Files the user picks, never URLs fetched. A task pane is a sandboxed
 * cross-origin iframe: fetching arbitrary image URLs is refused by CORS for
 * most hosts, and it would mean the add-in making a network request per row out
 * of somebody's data. A file picker needs no network, works offline, and keeps
 * the rule that values do not leave the pane.
 *
 * Shown only once the data actually refers to pictures, so a text-only merge
 * never sees it.
 */
function imageControl(doc: Document, state: PaneState): HTMLElement {
  const wrap = el(doc, "div", { class: "field images" });
  const tally = imageTally(state);
  const columns = imageColumns(state);

  wrap.append(
    el(doc, "span", {
      class: "caption",
      text: `${plural(tally.wanted, "picture")} named in ${columns.join(", ")}`,
    }),
  );
  wrap.append(
    el(doc, "input", {
      attrs: {
        type: "file",
        multiple: "",
        accept: "image/png,image/jpeg,image/gif,image/bmp",
        "data-field": "images",
      },
    }),
  );

  if (!state.images || state.images.size === 0) {
    wrap.append(
      el(doc, "p", {
        class: "muted",
        // Says what happens if they skip it, because skipping is allowed. A
        // merge with no pictures is not blocked — the frames keep their
        // placeholders, which is the same rule an unmatched text field follows.
        text: "Choose them from the folder they are in. Without them the merged slides keep the placeholder.",
      }),
    );
    return wrap;
  }

  wrap.append(
    el(doc, "p", {
      class: tally.missing.length > 0 ? "blocked" : "muted",
      text:
        tally.missing.length === 0
          ? `All ${plural(tally.wanted, "picture")} matched.`
          : `${tally.matched} of ${tally.wanted} matched. Missing: ${tally.missing.slice(0, 6).join(", ")}${tally.missing.length > 6 ? `, and ${tally.missing.length - 6} more` : ""}.`,
    }),
  );
  // Named once, and not as a problem. Picking a whole folder is the ordinary
  // way to do this, so most of the files being unused is expected.
  if (tally.spare.length > 0) {
    wrap.append(
      el(doc, "p", {
        class: "muted",
        text: `${plural(tally.spare.length, "file")} no row refers to — ignored.`,
      }),
    );
  }
  return wrap;
}

/**
 * A button per column, which puts `{{Column}}` where the cursor is.
 *
 * The question this step exists to answer, asked in those words by the first
 * person to use the add-in: "how do I insert the fields?" The answer used to be
 * "type them", which is a fine answer for one field and a bad one for twelve —
 * and it asks the user to spell a column name exactly right, from memory, with
 * the data in another window.
 *
 * COLUMNS, not fields. The list is what the data has, in the order it has it,
 * because that is the set the merge can fill: offering anything else would put
 * a placeholder on a slide that nothing will ever bind to. A column already on
 * the slides is marked rather than dropped — a template can legitimately use
 * one twice, and a list that shrinks as you press it loses its own order.
 *
 * The instruction line is not decoration. Nothing in Office.js reports where
 * the cursor is, so the pane cannot know whether the next press will land, and
 * `setSelectedDataAsync` on a slide with nothing selected refuses with a
 * message about a selection that reads as a bug. Saying "click into a text box
 * first" ahead of the press is the only version of that the user can act on.
 */
function insertControl(doc: Document, state: PaneState): HTMLElement {
  const wrap = el(doc, "div", { class: "insert" });
  // The engine's own reader decides which columns can carry a field, so the
  // button and the read-back cannot disagree. They did, for an hour, and it
  // shipped: see `insertableColumns`.
  const { can: columns, cannot } = insertableColumns(state);
  if (columns.length === 0 && cannot.length === 0) {
    // Reachable: `blockedReason` sends the user back for data before this step,
    // so the only way here with no columns is the Back link from Preview after
    // clearing the paste box. Says which, rather than rendering an empty list.
    wrap.append(el(doc, "p", { class: "muted", text: "Attach your data first — a field is a column name." }));
    return wrap;
  }
  wrap.append(
    el(doc, "p", {
      class: "muted",
      text: "Click into a text box on the slide, then press a column to put its field there.",
    }),
  );
  const list = el(doc, "ul", { class: "chips" });
  const placed = new Set(state.fields);
  const images = new Set(imageColumns(state));
  for (const column of columns) {
    const kind = images.has(column) ? "image" : undefined;
    const token = fieldToken(column, kind);
    const item = el(doc, "li");
    item.append(
      el(doc, "button", {
        class: "chip",
        text: token,
        attrs: {
          "data-insert": column,
          // Marked, not removed. See above.
          "data-placed": placed.has(column) ? "yes" : "no",
          // Which chips draw a PICTURE, so the two kinds are told apart before
          // they are pressed rather than after a merge.
          "data-kind": kind ?? "text",
          // The chip shows the TOKEN, which is what lands on the slide; the
          // title says what it is for. A user reading `{{Region}}` on a button
          // knows what will appear and not what it will become.
          title: kind
            ? `Draw a rectangle where the picture goes, then press this to put ${token} in it`
            : `Put ${token} on the slide${placed.has(column) ? " — already on these slides" : ""}`,
          ...(state.running ? { disabled: "" } : {}),
        },
      }),
    );
    list.append(item);
  }
  wrap.append(list);
  // Named, because a column nobody used is not an error and the user should not
  // have to diff two lists to see it. Only once something IS placed: before
  // that every column is unused and the line says nothing.
  const unused = unusedColumns(state).filter((c) => columns.includes(c));
  if (unused.length > 0 && state.fields.length > 0) {
    wrap.append(el(doc, "p", { class: "muted", text: `Not on a slide yet: ${unused.join(", ")}.` }));
  }
  // Named, not silently dropped. The fix is to rename the column, and a chip
  // that is simply absent says nothing about which one or why.
  if (cannot.length > 0) {
    wrap.append(
      el(doc, "p", {
        class: "blocked",
        text: `${cannot.join(", ")} ${cannot.length === 1 ? "cannot be a field" : "cannot be fields"}: a field name may not contain a brace or a pipe. Rename the ${cannot.length === 1 ? "column" : "columns"} and paste again.`,
      }),
    );
  }
  return wrap;
}

/**
 * Which slides are conditional, and on what.
 *
 * CLOSED by default and summarised shut, the same shape as the row picker and
 * for the same reason: most merges want every slide for every row, and a user
 * who never opens this should pay one line for it. The line states the current
 * answer rather than naming the feature, so it is discoverable without being in
 * the way.
 *
 * A SELECT of the columns, never free text. The engine matches a condition
 * against a column name exactly, so a typed name is a silent no-op the user
 * discovers by counting slides in the output — and the columns are already
 * known here, which is the whole reason this control lives on step 2 rather
 * than step 1 where the slide numbers come from.
 *
 * It does not remove the `unknownConditions` case and is not meant to: a
 * condition is chosen from THIS paste's columns, and the next paste may not
 * have them. That is what `danglingConditions` says before the merge and what
 * the engine reports after it.
 */
function conditionPicker(doc: Document, state: PaneState): HTMLElement {
  const slides = blockSlides(state);
  const wrap = el(doc, "div", { class: "conditions" });
  const set = slides.filter((n) => conditionFor(state, n) !== "").length;

  wrap.append(
    el(doc, "button", {
      class: "back conditions-toggle",
      text: state.conditionsOpen
        ? "Hide the conditions"
        : set === 0
          ? "Every slide, every row — add a condition"
          : `${plural(set, "slide")} conditional — change`,
      attrs: { "data-action": "conditions" },
    }),
  );
  if (!state.conditionsOpen) return wrap;

  wrap.append(
    el(doc, "p", {
      class: "muted",
      text: "A conditional slide is left out for a row whose cell in that column is empty, or reads false, no or 0.",
    }),
  );

  const list = el(doc, "ul", { class: "conditionlist" });
  for (const slide of slides) {
    const item = el(doc, "li");
    const label = el(doc, "label");
    label.append(el(doc, "span", { class: "caption", text: `Slide ${slide}` }));
    const select = el(doc, "select", { attrs: { "data-condition": String(slide) } });
    const chosen = conditionFor(state, slide);
    // "Always" first and selected by default, so the control opens saying what
    // is true rather than proposing the first column as an answer.
    select.append(el(doc, "option", { text: "Always", value: "" }));
    for (const column of state.columns ?? []) {
      select.append(el(doc, "option", { text: `Only when ${column}`, value: column }));
    }
    // A column the data no longer has. Kept as an option rather than dropped,
    // because dropping it would silently rewrite the user's answer to "Always"
    // and change what the merge produces with nothing said.
    if (chosen !== "" && !(state.columns ?? []).includes(chosen)) {
      select.append(el(doc, "option", { text: `Only when ${chosen} (no such column)`, value: chosen }));
    }
    // The PROPERTY, not the attribute — the pane rebuilds itself on every
    // change, and `selected` as an attribute is the default a control reverts
    // to. Same reason the row checkboxes set `.checked`.
    select.value = chosen;
    label.append(select);
    item.append(label);
    list.append(item);
  }
  wrap.append(list);

  const dangling = danglingConditions(state);
  if (dangling.length > 0) {
    wrap.append(
      el(doc, "p", {
        class: "blocked",
        // Says what will HAPPEN, not just that something is wrong. The engine
        // emits the slide anyway, and a user who expects it to be left out
        // would otherwise find out by counting slides in the output.
        text: `No column for ${dangling.join(", ")} — those slides will be included for every row.`,
      }),
    );
  }
  return wrap;
}

/**
 * Which rows to merge.
 *
 * CLOSED by default, and that is the whole shape of the control: 240 rows is
 * not a screen, and the overwhelmingly common case is "merge all of them". A
 * user who never opens it should not pay a single pixel for it, so shut it
 * reads as one line of summary and a link.
 *
 * A checkbox list with a search box, which is what the backlog specified and
 * what usage should get a chance to argue with before anything cleverer is
 * built. An expression language on a 320-pixel pane is a v2 question.
 *
 * The list is capped: a search that matches two hundred rows renders two
 * hundred checkboxes, and at that point the pane is a scroll bar. What is over
 * the cap is COUNTED rather than silently dropped — a list that stops without
 * saying so is one the user believes they have read.
 */
const ROW_LIST_CAP = 60;

function rowPicker(doc: Document, state: PaneState): HTMLElement {
  const records = state.records;
  const wrap = el(doc, "div", { class: "rows" });
  const total = records ? records.rows.length : 0;
  const taken = total - includedCount(state);

  const summary = el(doc, "button", {
    class: "back rows-toggle",
    text: state.rowsOpen
      ? "Hide the rows"
      : taken === 0
        ? `All ${plural(total, "row")} — choose which`
        : `${plural(taken, "row")} taken out — choose which`,
    attrs: { "data-action": "rows" },
  });
  wrap.append(summary);
  if (!state.rowsOpen || !records) return wrap;

  const label = el(doc, "label", { class: "field" });
  label.append(el(doc, "span", { class: "caption", text: "Search the rows" }));
  label.append(
    el(doc, "input", {
      value: state.rowSearch ?? "",
      attrs: { type: "text", autocomplete: "off", "data-field": "rowSearch" },
    }),
  );
  wrap.append(label);

  const matches = visibleRows(records, state.rowSearch ?? "");
  const shown = matches.slice(0, ROW_LIST_CAP);
  const list = el(doc, "ul", { class: "rowlist" });
  for (const index of shown) {
    const item = el(doc, "li");
    const line = el(doc, "label");
    const box = el(doc, "input", { attrs: { type: "checkbox", "data-row": String(index) } });
    // The PROPERTY, not the attribute — same reason the two slide-number boxes
    // use it: `checked` as an attribute is the default a control reverts to,
    // and this pane rebuilds itself on every change.
    box.checked = rowIncluded(state, index);
    line.append(box);
    line.append(el(doc, "span", { text: rowLabel(records, index) }));
    item.append(line);
    list.append(item);
  }
  wrap.append(list);

  if (matches.length === 0) {
    wrap.append(el(doc, "p", { class: "muted", text: "No row matches that." }));
  } else if (matches.length > shown.length) {
    // Counted, never silently dropped.
    wrap.append(
      el(doc, "p", {
        class: "muted",
        text: `Showing ${shown.length} of ${plural(matches.length, "match", "matches")} — search to narrow it.`,
      }),
    );
  }
  return wrap;
}
