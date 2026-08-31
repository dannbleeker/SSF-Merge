/**
 * Put the deck and the pane back to how a stranger would find them.
 *
 * A second capture run is not the same as a first one. The pane remembers the
 * last merge, so step 1 greets you with "Slides 2 to 3 repeat together, 3
 * times." and a "What this run did, step by step" disclosure — both true, both
 * impossible on a deck nobody has merged yet, and both in the photograph. The
 * shots from an un-reset run look fine and quietly advertise a state the
 * customer will not be in.
 *
 * Two separate things have to be undone, and the deck has to go first: the
 * pane's own "Remove these slides" is what knows which slides the merge added,
 * and reloading the pane throws that knowledge away.
 */
import { PANE, click, controls, currentSlide, evalIn, says, sleep, until } from "./pane.mjs";

const targets = async () =>
  (await (await fetch(process.env.SSF_CDP ?? "http://127.0.0.1:9333/json/list")).json()).filter(
    (t) => t.type === "page" || t.type === "iframe",
  );

/**
 * Undo the merge, if this pane is still showing one it can undo.
 *
 * "No merge to undo" is not the same as "the deck is clean", and the caller
 * checks the slide count afterwards for exactly that reason. The pane can lose
 * the offer while the extra slides stay: the add-in is served from GitHub
 * Pages, so merging a pull request redeploys it, the open pane reloads itself,
 * and the knowledge of which slides this merge added goes with it. That is not
 * hypothetical — it happened between a capture run and the next reset, and the
 * only visible symptom was the pane's build id changing.
 *
 * There are also TWO take-backs, not one. "Remove these slides" clears the six
 * that step 5 adds; "Remove the preview" clears the two that step 4 adds. A
 * capture run that stops at step 3 leaves the second kind, and a version of
 * this that only knew the first reported "no merge to undo" about a deck
 * holding five slides.
 */
async function unmerge() {
  const takeBacks = ["Remove these slides", "Remove the preview"];
  const offered = await controls();
  const found = takeBacks.find((label) => offered.some((c) => c.includes(label)));
  if (!found) {
    console.log("  nothing to take back (the pane is not offering it)");
    return;
  }
  await click(found);
  await until("Back to", { timeout: 120_000 });
  await sleep(2500);
  console.log(`  used "${found}"`);
}

/**
 * Drop the crumb the pane leaves in `localStorage` under `ssf-merge.run.v1`.
 *
 * Separate from the step the pane is on, and it survives the reload that clears
 * that. When a merge happens and the pane goes away before the take-back is
 * used, the next open greets you with "A merge from <date> added 6 slide(s) and
 * the pane closed before you could take them back."
 *
 * Which is the right thing for the product to say, and wrong in a photograph
 * for a store listing: it is a state a new customer cannot be in, and it dates
 * the screenshot. It appeared in a retake, above step 1, having been left by a
 * merge two days earlier — so the deck was clean, the step was right, and the
 * shot would still have been wrong.
 */
async function clearCrumb() {
  const had = await evalIn(
    PANE,
    `(() => {
       const key = 'ssf-merge.run.v1';
       const found = localStorage.getItem(key) !== null;
       localStorage.removeItem(key);
       return found;
     })()`,
  );
  console.log(had ? "  cleared the take-back crumb" : "  no take-back crumb to clear");
}

/**
 * Reload the add-in's own frame, which is what drops the step it is on.
 *
 * The pane keeps the current step in memory, not in `localStorage` — the crumb
 * under `ssf-merge.run.v1` is a different thing and is usually absent — so a
 * genuine reload is the whole reset.
 *
 * Proved by a marker rather than announced. The first version of this sent
 * `Page.reload` to the pane's CDP target, slept, and printed "reloaded the
 * pane". The frame ignored it completely: `Page.reload` needs its domain
 * enabled, and on this out-of-process target it silently did nothing at all.
 * The message printed anyway, for two runs, while the pane sat on step 5.
 */
async function reloadPane() {
  const target = (await targets()).find((t) => t.url.includes(PANE));
  if (!target) throw new Error("the add-in is not loaded; sideload it before resetting");

  await evalIn(PANE, `window.__ssfResetMarker = 'before'`);
  await evalIn(PANE, `location.reload()`);

  const deadline = Date.now() + 30_000;
  while (Date.now() < deadline) {
    await sleep(1000);
    // A reload tears the context down, so this throws for a moment first.
    const marker = await evalIn(PANE, `window.__ssfResetMarker ?? 'gone'`).catch(() => null);
    if (marker === "gone") {
      console.log("  reloaded the pane");
      await sleep(3000);
      return;
    }
  }
  throw new Error("the pane never reloaded: the marker set before it survived");
}

await unmerge();
// Before the reload, not after: the crumb is read as the pane boots, so
// clearing it afterwards leaves the notice on screen until the next reload.
await clearCrumb();
await reloadPane();

// Say what was actually achieved rather than "done". A reset that half-worked
// is the one failure this script can cause, and it shows up three shots later.
const deck = await currentSlide();
const pane = await says();
console.log(`  deck holds ${deck.of} slides`);
console.log(`  pane says: ${pane.slice(0, 120)}`);
if (deck.of !== 3) {
  // Name the remedy, not just the condition. Undoing a merge the pane has
  // forgotten would mean deleting slides through the rail on a guess about
  // which ones, and guessing wrong here quietly rewrites the demo deck. Opening
  // a fresh copy costs a minute and cannot be wrong.
  throw new Error(
    `the demo deck should be back to 3 slides and it holds ${deck.of}. ` +
      `The pane can only undo a merge it still remembers, and it forgets on reload — ` +
      `a Pages redeploy does that. Upload a fresh copy of ` +
      `docs/listing/demo/Quarterly-business-review.pptx and open that instead.`,
  );
}
if (!pane.includes("STEP 1")) throw new Error(`the pane should be on step 1; it says ${pane.slice(0, 160)}`);
if (/\d+ times|What this run did/.test(pane)) {
  throw new Error(`the pane is still remembering a previous run: ${pane.slice(0, 160)}`);
}
