/**
 * Walk the pane once and photograph every step, for the Marketplace listing.
 *
 * Preconditions, none of which this script can create for you:
 *   - Edge running against `test-kit/out/browser-profile` with CDP on 9333.
 *   - A signed-in PowerPoint for the web with `demo/Quarterly-business-review.pptx`
 *     open from a folder with a plain name. The file name is in every shot's
 *     title bar, and signing in is never the agent's to do.
 *   - The add-in sideloaded: Add-ins ▸ My Add-ins ▸ Upload My Add-in, pointed
 *     at `manifest-prod.xml`. That dialog is the one thing here that resisted
 *     automation completely: while it is open the host page reports a 0x0
 *     viewport so it cannot be photographed, and its tab strip lives in no CDP
 *     target that can be reached, so the two clicks have to be made by hand.
 *
 * Run it from the repository root:
 *
 *     node test-kit/driver/listing-shots.mjs
 *
 * It asserts its way through rather than sleeping and hoping. Every step waits
 * for the pane to say the thing that means the step is done, and the run stops
 * at the first thing that does not match, naming what it wanted and what the
 * pane actually said. A capture session that half-worked and wrote five files
 * anyway is worse than one that stopped, because the files look finished.
 */
import { readFileSync } from "node:fs";
import { click, controls, currentSlide, fill, says, selectSlides, until } from "./pane.mjs";
import { shoot } from "./shot.mjs";

const DEMO = "docs/listing/demo";
const SHOTS = "docs/listing/shots";

/** Slides 2 and 3 are the template block; slide 1 is the deck's cover. */
const BLOCK = [2, 3];

/** Press a button and fail loudly if it was missing or greyed out. */
async function press(label) {
  const got = await click(label);
  if (got !== "clicked") {
    throw new Error(
      `could not press ${JSON.stringify(label)}: ${got}. Pane offers: ${JSON.stringify(await controls())}`,
    );
  }
}

async function step(n, name, body) {
  console.log(`\n[${n}] ${name}`);
  await body();
}

await step(1, "mark the template block", async () => {
  await until("STEP 1");
  // Selected in the rail as well as typed into the fields. The numbers alone
  // would drive the merge, but the picture is about which slides repeat, and
  // that is a thing you point at rather than a thing you type.
  await selectSlides(BLOCK);
  await fill("First slide", String(BLOCK[0]));
  await fill("Last slide", String(BLOCK[1]));
  await until(`Slides ${BLOCK[0]} to ${BLOCK[1]}`);
  await shoot(`${SHOTS}/1-mark-the-block.png`);
});

await step(2, "attach the rows", async () => {
  await press(`Use slides ${BLOCK[0]} to ${BLOCK[1]}`);
  await until("STEP 2");
  // The same file `test/listing-demo.test.ts` merges, so the numbers in the
  // pictures and the numbers the test asserts cannot drift apart.
  await fill("Paste your rows, headers included", readFileSync(`${DEMO}/rows.txt`, "utf8").trimEnd());
  await until("3 rows attached");
  await shoot(`${SHOTS}/2-attach-your-rows.png`);
});

await step(3, "preview one row", async () => {
  await press("Use 3 rows");
  await until("STEP 3");
  await press("Use 4 fields");
  await until("STEP 4");
  await press("Preview the first row");
  await until("The first row is in your deck", { timeout: 90_000 });

  // The preview lands at the end of the deck, so the canvas is still showing
  // the template. Without this the shot is a card about a slide you cannot see.
  await selectSlides([4]);
  const at = await currentSlide();
  if (at.slide !== 4) throw new Error(`wanted the canvas on slide 4, it is on ${at.slide}`);
  await shoot(`${SHOTS}/3-preview-one-row.png`);
});

await step(4, "the sentence before the press", async () => {
  await press("On to the merge");
  await until("STEP 5");
  await until("6 slides will be added after slide 3, leaving 9 slides in the deck.");
  await shoot(`${SHOTS}/4-before-the-merge.png`);
});

await step(5, "after the merge", async () => {
  await press("Add 6 slides");
  // The merge is the slowest thing here by a wide margin: it is 27 placeholder
  // fills across six new slides, one Office round trip at a time.
  await until("6 slides added after slide 3", { timeout: 180_000 });
  await until("Remove slides 4 to 9, which this merge added.");
  await selectSlides([4]);
  const at = await currentSlide();
  if (at.of !== 9) throw new Error(`the deck should hold 9 slides, it holds ${at.of}`);
  await shoot(`${SHOTS}/5-after-the-merge.png`);
});

console.log(`\nfive shots in ${SHOTS}. The pane finished on:\n  ${(await says()).slice(0, 200)}`);
