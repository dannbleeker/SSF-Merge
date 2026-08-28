#!/usr/bin/env node
/**
 * Render every pane state to a PNG, at both ends of the width it must hold.
 *
 * The pane is the one surface the suite cannot judge. `pane-render.test.ts`
 * pins its behaviour in jsdom, which has no layout and no colour, so a rule
 * like the ORANGE BUDGET — one orange element per view — is invisible there
 * unless somebody thought to assert it. The first run of this script found
 * exactly that: the fields step drew the tick AND an orange-bordered chip.
 *
 *   npx vite --port 5199 --strictPort &
 *   node scripts/pane-shots.mjs
 *
 * 320 and 512 are the ends of the range a task pane is dragged between. A pane
 * judged at one width is a layout that breaks at the other.
 */
import { mkdirSync } from "node:fs";
import { chromium } from "playwright";

const PORT = process.env.PANE_PORT ?? "5199";
const OUT = process.env.PANE_SHOTS ?? "/tmp/pane-shots";

const full = {
  block: { from: 4, to: 6 },
  fields: ["First", "Last"],
  columns: ["First", "Last"],
  rows: 240,
  previewing: false,
  deckSize: 12,
};

const PASTE = "First\tLast\tEmail\nAda\tLovelace\tada@example.com\nGrace\tHopper\tgrace@example.com";

/**
 * What an Excel pivot table actually hands you, spaces and all.
 *
 * These three headers are the ones a real run pasted in, and the reader could
 * not see any of them: a field name was a list of allowed characters and a
 * space was not on it. Kept as a fixture because the screen that reported it —
 * the chips on the slides, "carry no fields yet" underneath — is the one nobody
 * was looking at.
 */
const PIVOT = ["Row Labels", "Min. of cost", "Sum of quantity monthly"];

/**
 * Data whose cells name pictures, which is what makes the picker appear.
 *
 * Pasted rather than described: the picker is drawn off `records`, which only a
 * real parse produces, so these fixtures hand the pane the same text a user
 * pastes and let `readPastedTable` decide — the same rule the "2-data" fixture
 * follows for its columns and row count.
 */
const PHOTO_PASTE = "First\tPhoto\nAda\tada.png\nGrace\tgrace.png\nAlan\talan.png";

const STATES = [
  { name: "1-template-empty", step: "template", state: { fields: [], previewing: false } },
  {
    name: "1-template-typing",
    step: "template",
    state: { fields: [], previewing: false, draft: { from: "4", to: "6" }, deckSize: 12 },
  },
  {
    name: "1-template-wrong-way-round",
    step: "template",
    state: { fields: [], previewing: false, draft: { from: "6", to: "4" }, deckSize: 12 },
  },
  { name: "1-template-chosen", step: "template", state: full },
  { name: "2-data-empty", step: "data", state: { ...full, columns: undefined, rows: undefined, paste: "" } },
  {
    // Self-consistent on purpose: the paste, the columns and the row count are
    // what `readPastedTable` answers for PASTE. A fixture whose label and whose
    // box disagree teaches the reader a bug that is not there.
    name: "2-data",
    step: "data",
    state: { ...full, paste: PASTE, columns: ["First", "Last", "Email"], rows: 2 },
  },
  {
    // The reported first run, on the step that now answers it: a fresh deck,
    // nothing on the slides, and a button per column. The old order refused at
    // step 1 and told the user to go and type names they had no way to know.
    name: "3-fields-nothing-placed",
    step: "fields",
    state: { ...full, fields: [], paste: PASTE, columns: ["First", "Last", "Email"], rows: 2 },
  },
  {
    name: "3-fields",
    step: "fields",
    state: {
      ...full,
      fields: ["First", "Last", "Nickname"],
      paste: PASTE,
      columns: ["First", "Last", "Email"],
      rows: 2,
    },
  },
  {
    // The reported screen, with the reader that can read it.
    name: "3-fields-pivot-headers",
    step: "fields",
    state: { ...full, fields: [], paste: PASTE, columns: PIVOT, rows: 2 },
  },
  {
    // A header the engine would read back as a DIFFERENT, shorter name. Named
    // rather than silently dropped: the fix is to rename the column.
    name: "3-fields-column-refused",
    step: "fields",
    state: { ...full, fields: [], paste: PASTE, columns: ["First", "Total|EUR"], rows: 2 },
  },
  {
    // The clipboard fallback, which is the outcome nobody sees unless they are
    // looking for it — an insert lands visibly on the slide, a copy lands
    // nowhere the user can see.
    name: "3-fields-clipboard",
    step: "fields",
    state: {
      ...full,
      fields: ["First"],
      paste: PASTE,
      columns: ["First", "Last", "Email"],
      rows: 2,
      fieldNote:
        "PowerPoint would not type it in, so {{Last}} is on your clipboard — click into a text box on the slide and paste it. (no insertion point)",
    },
  },
  {
    name: "4-preview-showing",
    step: "preview",
    state: { ...full, previewing: true, previewSlides: { from: 13, to: 15 } },
  },
  { name: "4-preview-idle", step: "preview", state: full },
  { name: "5-merge", step: "merge", state: full },
  { name: "5-merge-blocked", step: "merge", state: { fields: [], previewing: false } },
  {
    name: "5-merge-host-said",
    step: "merge",
    state: { ...full, notice: "PowerPoint would not name every slide between 4 and 6." },
  },
  // The states an adversarial review found nothing was rendering: a host call
  // in flight, a run that has landed, and a template missing two columns.
  {
    // With the draft, because in the real flow the boxes always hold what the
    // press was made from. A fixture whose boxes are empty while the button
    // reads "Reading the slides…" teaches a state that cannot happen.
    name: "1-template-reading",
    step: "template",
    state: { ...full, draft: { from: "4", to: "6" }, running: "inspect" },
  },
  { name: "5-merge-running", step: "merge", state: { ...full, running: "merge" } },
  {
    // Naming the call, which is the difference between a slow step and a stuck
    // one on a pane that is legitimately silent for two and a half minutes.
    name: "5-merge-waiting-on-host",
    step: "merge",
    state: { ...full, running: "merge", inFlight: "inserting the merged deck" },
  },
  {
    // The run record, collapsed. It must not bury the sentence above it.
    // The way back, which nothing rendered until 2026-08-27.
    // A run the pane never got to finish, offered back on the next open. The
    // slides are in the deck and the numbers were the only thing missing.
    name: "5-merge-recovered",
    step: "merge",
    state: {
      ...full,
      added: 720,
      deckSize: 732,
      notice: "A merge from 2026-08-27 added 720 slide(s) and the pane closed before you could take them back.",
    },
  },
  {
    name: "5-merge-done-undo",
    step: "merge",
    state: { ...full, added: 720, deckSize: 732, notice: "720 slides added after slide 12 · 480 placeholders filled." },
  },
  {
    name: "5-merge-done-with-log",
    step: "merge",
    state: {
      ...full,
      added: 720,
      deckSize: 732,
      notice: "720 slides added after slide 12 · no placeholders were filled — check the spelling in your template.",
      log: [
        "   0.0s  host  issued    call=counting the deck's slides budget=15000",
        "   0.1s  host  answered  call=counting the deck's slides ms=94",
        "   0.1s  host  issued    call=exporting the template slides budget=90000",
        "   2.4s  host  answered  call=exporting the template slides ms=2311",
        "   2.4s  host  issued    call=inserting the merged deck budget=60000",
        "  14.9s  host  answered  call=inserting the merged deck ms=12470",
      ].join("\n"),
    },
  },
  {
    // `deckSize` is the deck AFTER the merge, and two fixtures here were missing
    // that: they kept `full`'s pre-merge 12 beside an `added` of 720, and the
    // undo card obediently rendered `Remove slides -707 to 12`. Nobody had a
    // product bug and somebody was going to spend a morning finding that out —
    // the trap this script's own comments name twice, sprung a third time.
    //
    // The impossible state IS worth a shot; it just is not this one. See
    // "4-merge-nothing-to-take-back" below, where the deck is genuinely too
    // small and the card is correctly not drawn.
    name: "5-merge-done",
    step: "merge",
    state: { ...full, added: 720, deckSize: 732, notice: "720 slides added after slide 12." },
  },
  {
    // The slides went back by hand, or with Ctrl+Z, between the merge and this
    // draw — which the crash crumb makes reachable, because it offers a run
    // back on the NEXT open of the pane. No undo card: `sweepPlan` would refuse
    // it, so offering it would be a promise the next press breaks.
    name: "5-merge-nothing-to-take-back",
    step: "merge",
    state: { ...full, added: 720, deckSize: 12, notice: "720 slides added after slide 12." },
  },
  {
    // The record while the host has not answered — the state a wedged run sits
    // in, which used to show the waiting line and nothing else.
    name: "5-merge-running-with-log",
    step: "merge",
    state: {
      ...full,
      running: "merge",
      inFlight: "inserting the merged deck",
      log: [
        "   0.0s  pane  run starting  build=225e8a5 platform=OfficeOnline floor=1.2 deck=12",
        "   0.0s  host  issued    call=counting the deck's slides budget=15000",
        "   0.1s  host  answered  call=counting the deck's slides ms=94",
        "   0.1s  host  issued    call=exporting the template slides budget=90000",
        "   2.4s  host  answered  call=exporting the template slides ms=2311",
        "   2.4s  host  issued    call=inserting the merged deck budget=60000",
      ].join("\n"),
    },
  },
  {
    // A placeholder the author put in a chart. Not merged, and — until this —
    // not mentioned either: it is not in `fields`, so the chips never showed it.
    name: "3-fields-chart-placeholder",
    step: "fields",
    state: { ...full, paste: PASTE, columns: ["First", "Last", "Email"], rows: 2, unmergeable: ["Region"] },
  },
  {
    // The condition control, shut and open, on the merge step it moved to when
    // the fields step became about putting placeholders onto slides. Shut it is
    // one line that states the current answer; open it is a select per template
    // slide.
    name: "5-conditions-shut",
    step: "merge",
    state: { ...full, paste: PASTE, columns: ["First", "Last", "Email"], rows: 2 },
  },
  {
    name: "5-conditions-open",
    step: "merge",
    state: {
      ...full,
      paste: PASTE,
      columns: ["First", "Last", "Email"],
      rows: 2,
      conditionsOpen: true,
      conditions: { 5: "Email" },
    },
  },
  {
    // A condition naming a column this paste does not have — reachable by
    // choosing one and then pasting different data. The option is kept rather
    // than silently rewritten to "Always".
    name: "5-conditions-dangling",
    step: "merge",
    state: {
      ...full,
      paste: PASTE,
      columns: ["First", "Last", "Email"],
      rows: 2,
      conditionsOpen: true,
      conditions: { 4: "Renewal", 6: "Email" },
    },
  },
  {
    name: "3-fields-two-missing",
    step: "fields",
    state: { ...full, fields: ["First", "Nickname", "Badge"], paste: PASTE, columns: ["First", "Last"], rows: 2 },
  },
  {
    // The picker, before anything is chosen. It says what skipping costs,
    // because skipping is allowed.
    name: "2-data-pictures-wanted",
    step: "data",
    state: { ...full, paste: PHOTO_PASTE },
    paste: PHOTO_PASTE,
  },
  {
    name: "2-data-pictures-matched",
    step: "data",
    state: { ...full, paste: PHOTO_PASTE },
    paste: PHOTO_PASTE,
    files: ["ada.png", "grace.png", "alan.png"],
  },
  {
    // The screen that has to be readable: two pictures short, and a folder full
    // of files no row refers to. The missing ones are NAMED — a count sends the
    // user back through the spreadsheet to work out which.
    name: "2-data-pictures-missing",
    step: "data",
    state: { ...full, paste: PHOTO_PASTE },
    paste: PHOTO_PASTE,
    files: ["ada.png", "logo.png", "banner.png"],
  },
  {
    name: "1-template-past-the-end",
    step: "template",
    state: { fields: [], previewing: false, draft: { from: "4", to: "99" }, deckSize: 12 },
  },
];

// The bundled browser and the installed playwright can disagree on build
// number in this environment, so the binary is named rather than discovered.
const EXECUTABLE = process.env.CHROMIUM ?? "/opt/pw-browsers/chromium-1194/chrome-linux/chrome";

mkdirSync(OUT, { recursive: true });
const browser = await chromium.launch({ executablePath: EXECUTABLE });
let taken = 0;
for (const width of [320, 512]) {
  for (const { name, step, state, paste, files } of STATES) {
    const page = await browser.newPage({ viewport: { width, height: 620 } });
    await page.goto(`http://localhost:${PORT}/taskpane.html`);
    await page.evaluate(
      async ({ state, step, paste, files }) => {
        const { render } = await import("/render.ts");
        const shown = { ...state };
        // Built HERE rather than passed in: a parse result and a Map do not
        // survive the trip into the page, and a fixture that described its own
        // columns would be a fixture that can disagree with the parser.
        if (paste) {
          const { readPastedTable } = await import("/steps.ts");
          Object.assign(shown, readPastedTable(paste));
        }
        if (files) shown.images = new Map(files.map((name) => [name, new Uint8Array([1])]));
        render(document.getElementById("pane"), shown, step);
      },
      { state, step, paste, files },
    );
    await page.screenshot({ path: `${OUT}/${width}-${name}.png` });
    await page.close();
    taken++;
  }
}
await browser.close();
console.log(`${taken} shots in ${OUT}`);
