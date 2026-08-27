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
  { name: "2-fields-empty", step: "fields", state: { ...full, columns: undefined, rows: undefined, paste: "" } },
  {
    // Self-consistent on purpose: the paste, the columns and the row count are
    // what `readPastedTable` answers for PASTE. A fixture whose label and whose
    // box disagree teaches the reader a bug that is not there.
    name: "2-fields",
    step: "fields",
    state: {
      ...full,
      fields: ["First", "Last", "Nickname"],
      paste: PASTE,
      columns: ["First", "Last", "Email"],
      rows: 2,
    },
  },
  { name: "3-preview", step: "preview", state: { ...full, previewing: true } },
  { name: "3-preview-idle", step: "preview", state: full },
  { name: "4-merge", step: "merge", state: full },
  { name: "4-merge-blocked", step: "merge", state: { fields: [], previewing: false } },
  {
    name: "4-merge-host-said",
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
  { name: "4-merge-running", step: "merge", state: { ...full, running: "merge" } },
  {
    name: "4-merge-done",
    step: "merge",
    state: { ...full, added: 720, notice: "720 slides added after slide 12." },
  },
  {
    name: "2-fields-two-missing",
    step: "fields",
    state: { ...full, fields: ["First", "Nickname", "Badge"], paste: PASTE, columns: ["First", "Last"], rows: 2 },
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
  for (const { name, step, state } of STATES) {
    const page = await browser.newPage({ viewport: { width, height: 620 } });
    await page.goto(`http://localhost:${PORT}/taskpane.html`);
    await page.evaluate(
      async ({ state, step }) => {
        const { render } = await import("/render.ts");
        render(document.getElementById("pane"), state, step);
      },
      { state, step },
    );
    await page.screenshot({ path: `${OUT}/${width}-${name}.png` });
    await page.close();
    taken++;
  }
}
await browser.close();
console.log(`${taken} shots in ${OUT}`);
