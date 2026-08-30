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
 * judged at one width is a layout that breaks at the other. Both THEMES, too:
 * PowerPoint can be dark while the OS is light, and the dark palette is a
 * separate set of tokens that nothing else in this repo exercises.
 *
 * It also MEASURES, because two of the things a screenshot shows are numbers a
 * reader cannot take off a PNG, and both have produced real defects:
 *
 * - **horizontal overflow.** A single unbroken column header took the pane to
 *   545px inside a 320px frame, and a spaceless error from PowerPoint to 3751 —
 *   with the one filled button off the side. Every long string on this screen
 *   comes from outside it.
 * - **text contrast.** Blue is a background here (the header, the primary
 *   button, both carrying white text) and it was also the ink on chips, field
 *   tags and every secondary button. On the dark palette that ink was 3.0:1,
 *   and "Remove these slides" — the whole way back from a merge — was 2.93:1.
 *
 * Findings are printed and the process exits 1, so this can be read by a person
 * or wired to something. Disabled controls are exempt from the contrast rule,
 * which is what WCAG 1.4.3 says and not a convenience: a greyed-out button is
 * meant to read as unavailable.
 */
import { existsSync, mkdirSync, readFileSync } from "node:fs";
import { chromium } from "playwright";

/**
 * axe-core, read off disk and injected per page.
 *
 * The fourth thing measured here, added 2026-08-30 after a sweep found the
 * renderer carrying exactly ONE `aria` attribute. It answers the half of
 * accessibility a person cannot eyeball — names, roles, labels, duplicate ids,
 * a control with nothing to call it — over the same state list, both widths and
 * both themes, which is 72 more renders than anybody was going to check by
 * hand.
 *
 * What it does NOT answer is the half that found the real defect: whether the
 * pane SAYS anything when it changes. Nothing static can see a missing live
 * region, and axe was green on the pane before the region existed. It is a
 * floor, not a verdict.
 */
const AXE = readFileSync("node_modules/axe-core/axe.min.js", "utf8");

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

/**
 * A header with nothing in it a line can break at.
 *
 * `overflow-wrap` is what keeps this inside a 320px frame, and nothing else in
 * these fixtures has a word long enough to need it — the pivot headers all
 * carry spaces. A system export writes names like this one.
 */
const UNBROKEN_PASTE = "First\tSumOfQuantityMonthlyForTheNorthernRegionIncludingSubsidiaries\nAda\t42";

/** Enough rows to reach the list's cap and its "search to narrow it" line. */
const MANY_ROWS = ["First", ...Array.from({ length: 200 }, (_, i) => `Person number ${i + 1} of two hundred`)].join(
  "\n",
);

/**
 * Two pictures the data tells apart and a file picker cannot.
 *
 * Matching is by base name, because `File.name` has no path in it, so these
 * are one name by the time anything can act on them — and the pane says so
 * rather than filling both from the same file in silence.
 */
const CLASH_PASTE = "First\tPhoto\nAda\tregions/eu/logo.png\nBo\tregions/us/logo.png";

/** Rows with a blank, which is what the blank-cell control is about. */
const BLANK_PASTE = [
  "First\tLast\tNotes",
  ...Array.from({ length: 24 }, (_, i) => `Person ${i + 1}\tSurname ${i + 1}\t${i % 4 === 0 ? "" : `note ${i + 1}`}`),
].join("\n");

/** The fields per slide of the three-slide block, as a template read answers. */
const BLANK_SLIDE_FIELDS = [["First"], ["Last"], ["Notes"]];

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
  /**
   * The two shapes the overflow half of the audit exists for.
   *
   * Without them that half is a measurement nothing exercises — a gate whose
   * name is wider than its fixtures — and it would report clean against a
   * stylesheet that had lost `overflow-wrap`. Both are real: a header with no
   * spaces in it is what a system export writes, and a host error with no
   * spaces is what `readable` caps at 400 characters. A cap is not a break.
   */
  {
    name: "2-data-unbroken-header",
    step: "data",
    state: { ...full, paste: UNBROKEN_PASTE },
    paste: UNBROKEN_PASTE,
  },
  /**
   * The row picker open, which nothing here drew.
   *
   * It is the densest thing the pane renders — a scrolling list of checkboxes,
   * a search box, a cap message — and none of it had ever been measured or
   * looked at.
   */
  {
    name: "5-rows-open",
    step: "merge",
    state: { ...full, paste: MANY_ROWS, rowsOpen: true, excluded: [1] },
    paste: MANY_ROWS,
  },
  {
    name: "5-rows-searched",
    step: "merge",
    state: { ...full, paste: MANY_ROWS, rowsOpen: true, rowSearch: "nothing matches this" },
    paste: MANY_ROWS,
  },
  /**
   * Every merge caution at once. Each is a different thing about to happen and
   * none may swallow another, so the screen has to hold all three.
   */
  {
    name: "5-merge-every-caution",
    step: "merge",
    state: {
      ...full,
      paste: CLASH_PASTE,
      fields: ["First", "Nickname", "Photo"],
      imageFields: [],
    },
    paste: CLASH_PASTE,
    files: ["logo.png"],
  },
  {
    name: "2-data-picture-clash",
    step: "data",
    state: { ...full, paste: CLASH_PASTE },
    paste: CLASH_PASTE,
    files: ["logo.png"],
  },
  /**
   * What a blank cell does, all three answers, plus the shape it exists for:
   * a blank in a field that only appears on a CONDITIONAL slide.
   */
  {
    name: "5-empties-shut",
    step: "merge",
    state: { ...full, paste: BLANK_PASTE, slideFields: BLANK_SLIDE_FIELDS },
    paste: BLANK_PASTE,
  },
  {
    name: "5-empties-open",
    step: "merge",
    state: { ...full, paste: BLANK_PASTE, slideFields: BLANK_SLIDE_FIELDS, emptiesOpen: true },
    paste: BLANK_PASTE,
  },
  {
    name: "5-empties-skip",
    step: "merge",
    state: {
      ...full,
      paste: BLANK_PASTE,
      slideFields: BLANK_SLIDE_FIELDS,
      onEmpty: "skip",
      emptiesOpen: true,
    },
    paste: BLANK_PASTE,
  },
  {
    // Shut, with the count above the button and both figures in the heading.
    name: "5-empties-skip-shut",
    step: "merge",
    state: { ...full, paste: BLANK_PASTE, slideFields: BLANK_SLIDE_FIELDS, onEmpty: "skip" },
    paste: BLANK_PASTE,
  },
  {
    name: "5-merge-spaceless-notice",
    step: "merge",
    state: {
      ...full,
      paste: PASTE,
      columns: ["First", "Last", "Email"],
      rows: 2,
      notice: `PowerPoint refused the insert: ${"GeneralException".repeat(20)}`,
    },
  },
];

// The bundled browser and the installed playwright can disagree on build
// number in THIS environment, so the binary is named rather than discovered.
//
// Only when it is actually there. On a CI runner playwright installs its own
// and knows where it put it, and naming a path that does not exist fails the
// launch with an error about a missing executable rather than about anything
// this script is for. `CHROMIUM` overrides both.
const CONTAINER_CHROMIUM = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome";
const EXECUTABLE = process.env.CHROMIUM || (existsSync(CONTAINER_CHROMIUM) ? CONTAINER_CHROMIUM : undefined);

/**
 * What the page can say about itself that a PNG cannot.
 *
 * Runs INSIDE the page, so it reads computed styles and real geometry rather
 * than anything this file believes about the stylesheet.
 *
 * Contrast is measured against the nearest ancestor with a non-transparent
 * background, which is what the eye sees; a colour with no opaque ground
 * behind it falls back to white rather than being skipped, because skipping is
 * how a measurement quietly stops measuring.
 */
function audit() {
  const numbers = (/** @type {string} */ s) => (s.match(/[\d.]+/g) ?? []).map(Number);
  const luminance = (/** @type {number[]} */ [r, g, b]) => {
    const channel = (v) => {
      const x = v / 255;
      return x <= 0.03928 ? x / 12.92 : ((x + 0.055) / 1.055) ** 2.4;
    };
    return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b);
  };
  const groundOf = (/** @type {Element} */ el) => {
    for (let node = el; node; node = node.parentElement) {
      const colour = numbers(getComputedStyle(node).backgroundColor);
      if (colour.length >= 3 && (colour[3] === undefined || colour[3] > 0)) return colour;
    }
    return [255, 255, 255];
  };

  const findings = [];
  const root = document.documentElement;
  // The DOCUMENT, not the pane: a child can overhang its parent harmlessly,
  // and what matters is whether the frame has to scroll sideways.
  if (root.scrollWidth > root.clientWidth) {
    const past = [];
    for (const el of document.querySelectorAll("#pane *")) {
      const box = el.getBoundingClientRect();
      if (box.right > root.clientWidth + 0.5 || box.left < -0.5) {
        past.push(`${el.tagName.toLowerCase()}.${el.className || "-"} "${(el.textContent ?? "").trim().slice(0, 40)}"`);
      }
    }
    // A block-level element's BOX is clamped to its container, so the one that
    // overflows is usually a paragraph whose text is wider than it is. Name
    // the widest by `scrollWidth` when no box is past the edge, or the finding
    // says only that something is wrong.
    if (past.length === 0) {
      let widest = null;
      for (const el of document.querySelectorAll("#pane *")) {
        // An element that scrolls or clips its own content cannot push the
        // document — a textarea's `scrollWidth` is its text, and naming it
        // sends the reader at the one box on the screen that is fine.
        if (getComputedStyle(el).overflowX !== "visible") continue;
        if (el.scrollWidth > root.clientWidth && (!widest || el.scrollWidth > widest.scrollWidth)) widest = el;
      }
      if (widest) {
        past.push(
          `${widest.tagName.toLowerCase()}.${widest.className || "-"} is ${widest.scrollWidth}px wide "${(widest.textContent ?? "").trim().slice(0, 40)}"`,
        );
      }
    }
    findings.push(`overflows sideways: ${root.scrollWidth}px in ${root.clientWidth} — ${past[0] ?? "nothing named"}`);
  }

  const walk = (el) => {
    for (const node of el.childNodes) {
      if (node.nodeType === 3 && (node.textContent ?? "").trim() !== "") {
        // WCAG 1.4.3 exempts an inactive control: greyed out is meant to read
        // as unavailable, and holding it to 4.5:1 would remove the signal.
        if (el.disabled === true || el.closest("[disabled]")) continue;
        const style = getComputedStyle(el);
        const ratio =
          (Math.max(luminance(numbers(style.color)), luminance(groundOf(el))) + 0.05) /
          (Math.min(luminance(numbers(style.color)), luminance(groundOf(el))) + 0.05);
        const size = Number.parseFloat(style.fontSize);
        const large = size >= 24 || (size >= 18.66 && Number.parseInt(style.fontWeight, 10) >= 700);
        const need = large ? 3 : 4.5;
        if (ratio < need) {
          findings.push(
            `contrast ${ratio.toFixed(2)}:1 (needs ${need}) on ${el.tagName.toLowerCase()}.${el.className || "-"} "${(node.textContent ?? "").trim().slice(0, 40)}"`,
          );
        }
      } else if (node.nodeType === 1) walk(node);
    }
  };
  walk(document.body);

  // WHERE THE KEYBOARD IS, which is the third thing a PNG cannot show and the
  // one a screenshot cannot show at all: a focus ring is only on screen while
  // something is focused.
  //
  // It found the sharpest of the three. Buttons were not in the stylesheet's
  // focus rule, so they fell back to Chrome's own ring — `rgb(16, 16, 16)`,
  // near-black, drawn on the dark theme's near-black pane at 1.03:1. Every
  // chip, both disclosures, the back links and the undo button, with no
  // visible focus anywhere in that theme except inside a text box.
  //
  // The caller presses Tab once before this runs. Chrome matches
  // `:focus-visible` on a programmatic `focus()` only once the last
  // interaction was a keyboard one, so without that every button here reports
  // no outline at all — a measurement that answers "clean" because it never
  // looked.
  const focusable = document.querySelectorAll(
    "#pane button:not([disabled]), #pane input:not([disabled]), #pane select:not([disabled]), #pane textarea:not([disabled])",
  );
  const already = new Set();
  for (const el of focusable) {
    el.focus();
    if (document.activeElement !== el) continue; // not actually focusable
    const style = getComputedStyle(el);
    const key = `${el.tagName}.${el.className}`;
    if (already.has(key)) continue;
    already.add(key);
    const width = Number.parseFloat(style.outlineWidth);
    if (!(width > 0) || style.outlineStyle === "none") {
      findings.push(`no focus ring on ${el.tagName.toLowerCase()}.${el.className || "-"}`);
      continue;
    }
    const ratio =
      (Math.max(luminance(numbers(style.outlineColor)), luminance(groundOf(el.parentElement ?? el))) + 0.05) /
      (Math.min(luminance(numbers(style.outlineColor)), luminance(groundOf(el.parentElement ?? el))) + 0.05);
    // 3:1, which is what WCAG 1.4.11 asks of a non-text indicator.
    if (ratio < 3) {
      findings.push(`focus ring ${ratio.toFixed(2)}:1 (needs 3) on ${el.tagName.toLowerCase()}.${el.className || "-"}`);
    }
  }
  return findings;
}

mkdirSync(OUT, { recursive: true });
const browser = await chromium.launch(EXECUTABLE ? { executablePath: EXECUTABLE } : {});
let taken = 0;
// Deduplicated across states: the same chip is on eight screens, and eight
// copies of one finding is a report nobody reads to the end.
const found = new Map();
for (const width of [320, 512]) {
  for (const theme of ["light", "dark"]) {
    for (const { name, step, state, paste, files } of STATES) {
      const page = await browser.newPage({ viewport: { width, height: 620 } });
      // Office.js is fetched from Microsoft by `taskpane.html` and is not what
      // is being measured here — `render` is called directly. Refused rather
      // than waited on: on a machine that cannot reach it, every one of these
      // pages otherwise spends its connect timeout before drawing anything.
      await page.route("https://appsforoffice.microsoft.com/**", (route) => route.abort());
      await page.goto(`http://localhost:${PORT}/taskpane.html`);
      await page.evaluate(
        async ({ state, step, paste, files, theme }) => {
          // The stamp `main.ts` writes from `Office.context.officeTheme`. Set
          // before the render so the first paint is the one measured.
          document.documentElement.setAttribute("data-theme", theme);
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
        { state, step, paste, files, theme },
      );
      await page.screenshot({ path: `${OUT}/${width}-${theme}-${name}.png` });
      // BEFORE the Tab below: a focus ring in the accessibility tree is not
      // what axe is being asked about, and the shot above is the clean state.
      await page.addScriptTag({ content: AXE });
      // The strings are built INSIDE the page and only strings come back —
      // axe's result object is a large `any` across the boundary, and mapping
      // it here rather than there is how this file would start carrying an
      // untyped shape it never uses.
      /** @type {string[]} */
      const violations = await page.evaluate(async () => {
        /** @type {{ violations: { impact: string; id: string; help: string }[] }} */
        const run = await globalThis.axe.run(document, { resultTypes: ["violations"] });
        return run.violations.map((x) => `${x.impact}: ${x.id} — ${x.help}`);
      });
      for (const v of violations) {
        if (!found.has(v)) found.set(v, `${width} ${theme} ${name}`);
      }
      // AFTER the shot, and before the audit. It tells Chrome the last
      // interaction was a keyboard one, which is what makes `:focus-visible`
      // match the programmatic `focus()` the focus sweep uses — and it would
      // put a ring in the picture if it ran first.
      await page.keyboard.press("Tab");
      for (const finding of await page.evaluate(audit)) {
        if (!found.has(finding)) found.set(finding, `${width} ${theme} ${name}`);
      }
      await page.close();
      taken++;
    }
  }
}
await browser.close();
console.log(`${taken} shots in ${OUT}`);
if (found.size === 0) {
  console.log(
    "audit: nothing overflows, every live label clears its contrast floor, every control shows where the keyboard is, " +
      "and axe finds no violation",
  );
} else {
  console.log(`audit: ${found.size} finding(s)`);
  for (const [finding, where] of found) console.log(`  ${where}: ${finding}`);
  process.exitCode = 1;
}
