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

const STATES = [
  { name: "1-template-empty", step: "template", state: { fields: [], previewing: false } },
  { name: "1-template-chosen", step: "template", state: full },
  { name: "2-fields", step: "fields", state: { ...full, fields: ["First", "Last", "Nickname"] } },
  { name: "3-preview", step: "preview", state: { ...full, previewing: true } },
  { name: "4-merge", step: "merge", state: full },
  { name: "4-merge-blocked", step: "merge", state: { fields: [], previewing: false } },
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
