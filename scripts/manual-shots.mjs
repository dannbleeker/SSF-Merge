#!/usr/bin/env node
/**
 * The pictures in `docs/MANUAL.md`, rendered from the real pane.
 *
 * Separate from `pane-shots.mjs` on purpose, and the split is by JOB rather
 * than by convenience. That script is an AUDIT: every state the pane can
 * reach, both widths, both themes, measured for overflow and contrast, exit 1
 * on a finding. This one is a DOCS BUILD: one picture per state in `STATES`
 * below, one width, one theme, committed to the repository so GitHub can
 * render them.
 *
 * The count is deliberately not written out here. It used to be, and it said
 * six while `STATES` held eight — a number in a comment beside the list it
 * counts is a number that drifts, and this file's whole subject is a picture
 * that disagrees with what it claims to show.
 *
 * Wiring the manual to the audit's output would have coupled a document to a
 * tool whose whole point is to grow — 172 shots today — and would have put the
 * manual's pictures behind a script that deliberately fails.
 *
 *   npx vite --port 5199 --strictPort &
 *   node scripts/manual-shots.mjs
 *
 * **The states tell ONE story, and the numbers across them agree.** A deck of 8
 * slides, a template on slide 3, three rows of three columns, three slides
 * added, eleven in the deck afterwards, nine placeholders filled — three
 * paragraphs a slide, three slides. Most of them walk that merge through in
 * order; the last two are the same deck going WRONG, an unmatched field at the
 * step where it is cheap to fix and at the step where it is not. They spread
 * the same fixture as the rest, so they are not a second story with numbers of
 * their own. `pane-shots.mjs` has a comment recording what it cost to get this
 * wrong twice there: a fixture pairing a pre-merge deck size with a post-merge
 * count drew `Remove slides -707 to 12`, and somebody was going to spend a
 * morning hunting a product bug that did not exist. In a MANUAL that is worse
 * than a wasted morning — the reader has no way to know the screenshot is
 * lying, and they will type the numbers back at you when they report it.
 *
 * One theme and one width, both chosen rather than defaulted. **Light**,
 * because it is what PowerPoint opens as, and a manual is not the place to
 * exercise a palette — that is what the audit's dark pass is for. **380px**,
 * because it sits between the 320 and 512 the audit measures: wide enough that
 * the text in the picture is legible on a phone reading GitHub, narrow enough
 * that it is honestly the shape of a task pane rather than a browser window.
 */
import { mkdirSync } from "node:fs";
import { chromium } from "playwright";

const PORT = process.env.PANE_PORT ?? "5199";
const OUT = process.env.MANUAL_SHOTS ?? "docs/images";
const EXECUTABLE = process.env.PLAYWRIGHT_CHROMIUM;

/**
 * ONE browser, named, because the pictures depend on which one draws them.
 *
 * A refresh on 2026-08-30 produced a `step-4-preview.png` twenty pixels shorter
 * than the committed one — same commit, same script, same words, the card's
 * sentence wrapping onto two lines instead of three. Every other shot in the
 * set matched to the pixel. Two local browsers, Edge and Chrome, agreed with
 * each other and disagreed with what was in the repository, so it was not this
 * machine's fonts and not the code: it was the binary. Which binary drew the
 * committed set had never been recorded, so the cause could not be established
 * afterwards — only the effect.
 *
 * That is worth removing rather than explaining. These images are committed and
 * read as diffs: a set that reflows depending on who ran it makes every future
 * refresh look like a change, and lets a real change hide in the noise. So the
 * channel is named here rather than being whatever Playwright happens to have
 * installed.
 *
 * `msedge` rather than the bundled Chromium for a second reason: AppLocker on
 * the machine this is developed on refuses binaries outside Program Files, and
 * Playwright's own Chromium lives under %LOCALAPPDATA%. It fails there with
 * `spawn UNKNOWN`, which reads like a broken script rather than a policy.
 *
 * `PLAYWRIGHT_CHROMIUM` still overrides, for a machine with no Edge. Expect the
 * wrapping to move if you use it, and do not commit the result as though
 * nothing had changed.
 *
 * **The committed pictures predate this pin**, so the first refresh under it
 * rewrites all eight: same size, same words, lighter text, and the headline
 * wrapping a word later. That is the binary changing, not the pane. It was left
 * for the owner to accept deliberately rather than arriving inside an unrelated
 * commit — a manual whose pictures all changed is a thing somebody should have
 * chosen. Once that refresh lands, a diff here means the pane moved.
 */
const CHANNEL = "msedge";

const WIDTH = 380;

/**
 * The walkthrough's data, and it is the same three rows the manual prints.
 *
 * Real tab characters, because that is what a paste out of Excel is and what
 * `parseDelimited` sniffs for first. Typed here with `\t` so the file survives
 * an editor that trims whitespace.
 */
const PASTE = [
  "First\tCity\tRole",
  "Ada\tLondon\tAnalyst",
  "Grace\tNew York\tEngineer",
  "Katherine\tHampton\tMathematician",
].join("\n");

/** The deck the walkthrough starts from, before anything is added. */
const DECK = 8;

/** What every screen after step 2 knows: the block, the columns, the rows. */
const attached = {
  block: { from: 3, to: 3 },
  columns: ["First", "City", "Role"],
  rows: 3,
  paste: PASTE,
  previewing: false,
  deckSize: DECK,
};

const STATES = [
  {
    name: "step-1-template",
    step: "template",
    // The boxes hold what the press will be made from. An empty pair beside a
    // live button is a state the pane cannot be in.
    state: { fields: [], previewing: false, draft: { from: "3", to: "3" }, deckSize: DECK },
  },
  {
    name: "step-2-data",
    step: "data",
    state: { ...attached, fields: [] },
  },
  {
    // Every field matched. The unmatched case has a screen of its own below,
    // because it is the commonest thing a first merge gets wrong and a manual
    // that only shows the happy path is no use on the day it goes wrong.
    name: "step-3-fields",
    step: "fields",
    state: { ...attached, fields: ["First", "City", "Role"] },
  },
  {
    name: "step-4-preview",
    step: "preview",
    // Showing rather than idle: the preview's own slide is the thing the step
    // is about, and the card that names it is what the reader needs to
    // recognise when they go looking for those slides in the rail.
    state: { ...attached, fields: ["First", "City", "Role"], previewing: true, previewSlides: { from: 9, to: 9 } },
  },
  {
    name: "step-5-merge",
    step: "merge",
    state: { ...attached, fields: ["First", "City", "Role"] },
  },
  {
    name: "step-5-done",
    step: "merge",
    state: {
      ...attached,
      fields: ["First", "City", "Role"],
      added: 3,
      // The deck AFTER the merge. Pairing this with the pre-merge 8 is the
      // exact mistake the header block above records.
      deckSize: DECK + 3,
      // `describeMerge`'s own shape: clauses joined with " · ", one sentence.
      // Nine is three paragraphs a slide over three slides, which is what the
      // walkthrough's template actually holds.
      // Without this the undo card is not drawn at all — `undoIsPossible`
      // takes all three numbers and `sweepPlan` refuses a plan it cannot floor
      // against the deck's size before the run. The walkthrough's last
      // paragraph is about that card, so a picture missing it would document
      // the one thing the step does not appear to offer.
      deckAtStart: DECK,
      notice: "3 slides added after slide 8 · 9 placeholders filled.",
    },
  },
  {
    // The same mistake one step earlier, where it is CHEAP to fix. `Roel` has
    // no column, so its chip is outlined and takes the screen's one orange —
    // which is the whole of how the fields step reports it, and is not a thing
    // a sentence in the manual can show.
    name: "step-3-fields-unmatched",
    step: "fields",
    state: { ...attached, fields: ["First", "City", "Roel"] },
  },
  {
    // The failure the walkthrough warns about, on the screen that reports it.
    // A misspelled placeholder is a WARNING here and not a refusal, and the
    // picture is the only way to show that the button is still live.
    name: "field-with-no-column",
    step: "merge",
    state: { ...attached, fields: ["First", "City", "Roel"] },
  },
];

mkdirSync(OUT, { recursive: true });
const browser = await chromium.launch(EXECUTABLE ? { executablePath: EXECUTABLE } : { channel: CHANNEL });
// Said out loud, and recorded in the run, because the last time these pictures
// disagreed nobody could say what had drawn the ones already committed.
console.log(`browser : ${EXECUTABLE ? `${EXECUTABLE} (PLAYWRIGHT_CHROMIUM)` : CHANNEL} — ${browser.version()}`);
for (const { name, step, state } of STATES) {
  const page = await browser.newPage({ viewport: { width: WIDTH, height: 640 } });
  // Office.js is fetched from Microsoft by `taskpane.html` and is not what is
  // being pictured — `render` is called directly. Refused rather than waited
  // on, so a machine that cannot reach it does not spend a connect timeout per
  // shot before drawing anything.
  await page.route("https://appsforoffice.microsoft.com/**", (route) => route.abort());
  await page.goto(`http://localhost:${PORT}/taskpane.html`);
  await page.evaluate(
    async ({ state, step }) => {
      document.documentElement.setAttribute("data-theme", "light");
      const { render } = await import("/render.ts");
      const shown = { ...state };
      // The parse is done HERE rather than passed in: a fixture that described
      // its own columns could disagree with the parser, and then the picture
      // shows a pane nobody can get to.
      if (shown.paste) {
        const { readPastedTable } = await import("/steps.ts");
        Object.assign(shown, readPastedTable(shown.paste));
      }
      render(document.getElementById("pane"), shown, step);
    },
    { state, step },
  );
  // Fit the viewport to what was drawn, then shoot the viewport.
  //
  // The two obvious ways are both wrong here, and each was tried. A plain
  // viewport shot crops the one filled button off the bottom of the taller
  // screens — the thing every one of these pictures exists to show. Shooting
  // the `#pane` element instead fits every screen exactly and loses the
  // `<header>`, which lives in `taskpane.html` beside `#pane` rather than
  // inside it: the pictures then start mid-pane, with no "SSF Merge" bar for a
  // reader to match against the thing on their screen.
  const height = await page.evaluate(() => document.body.scrollHeight);
  await page.setViewportSize({ width: WIDTH, height });
  await page.screenshot({ path: `${OUT}/${name}.png` });
  await page.close();
}
await browser.close();
console.log(`${STATES.length} manual shots in ${OUT}`);
