/**
 * The Office.js calls, and nothing else.
 *
 * Every judgement this file needs is imported from `src/host`, where it is a
 * pure function the suite can check: which version floor the host clears, where
 * the deck's bytes come from, whether an insert did what it said, and which
 * slides an undo may remove. Nothing here decides anything, because a decision
 * inside a `PowerPoint.run` callback is a decision nobody can test and the host
 * this add-in runs on is documented to lie about ids, to accept calls it does
 * not perform, and to answer differently on two runs of the same build.
 *
 * The rules it obeys, all of them measured rather than assumed — see
 * `CLAUDE.md`:
 *
 * - an insert always names a `targetSlideId`, or the host puts the new slides
 *   in FRONT of the user's title slide;
 * - undo is positional and clamped, never by id;
 * - a queued delete that raised nothing has not necessarily happened, so the
 *   deck is counted again afterwards;
 * - the deck DELTA is the evidence, never the absence of an error: an insert
 *   has timed out having landed everything it was asked for;
 * - every call is bounded, because a call that stops answering here never
 *   comes back.
 */
import { chooseDeckSource, checkFloor, templateOffset, type DeckSource, type Supports } from "../host/capability.js";
import { insertVerdict, type InsertVerdict } from "../host/verdicts.js";
import { sweepPlan } from "../host/undo.js";
import { BUDGET, withTimeout } from "../host/timeout.js";

/** What the host says it supports, as the pure layer wants it. */
export const hostSupports: Supports = (version) => Office.context.requirements.isSetSupported("PowerPointApi", version);

/** Whether this host can run the add-in at all. Ask before anything else. */
export function ready(): { ok: boolean; detail: string } {
  return checkFloor(hostSupports);
}

export async function slideCount(): Promise<number> {
  return withTimeout(
    PowerPoint.run(async (context) => {
      const count = context.presentation.slides.getCount();
      await context.sync();
      return count.value;
    }),
    BUDGET.read,
    "counting the deck's slides",
  );
}

/** The id of the last slide, which is what an insert is anchored after. */
export async function lastSlideId(): Promise<string> {
  const before = await slideCount();
  return withTimeout(
    PowerPoint.run(async (context) => {
      const last = context.presentation.slides.getItemAt(before - 1);
      last.load("id");
      await context.sync();
      return last.id;
    }),
    BUDGET.read,
    "reading the last slide's id",
  );
}

export interface TemplateBytes {
  /** The package, base64. */
  base64: string;
  /** Index of the first template slide INSIDE that package — not in the deck. */
  offset: number;
  source: DeckSource;
  detail: string;
}

/**
 * The template block's bytes.
 *
 * `slideIds` are the template slides in deck order and `blockStartInDeck` is
 * where the first of them sits, because the two routes return packages of
 * different shapes and only the caller knows which slides it wants.
 * `templateOffset` is what reconciles them.
 */
export async function readTemplate(slideIds: string[], blockStartInDeck: number): Promise<TemplateBytes> {
  const choice = chooseDeckSource(hostSupports);
  const offset = templateOffset(choice.source, blockStartInDeck);
  if (choice.source === "subset") {
    const base64 = await withTimeout(
      PowerPoint.run(async (context) => {
        const bytes = context.presentation.slides.exportAsBase64Presentation(slideIds);
        await context.sync();
        return bytes.value;
      }),
      BUDGET.file,
      "exporting the template slides",
    );
    return { base64, offset, source: choice.source, detail: choice.detail };
  }
  const base64 = await withTimeout(readDeckThroughFileApi(), BUDGET.file, "reading the presentation's bytes");
  return { base64, offset, source: choice.source, detail: choice.detail };
}

/**
 * The whole package, slice by slice.
 *
 * The floor every host has; the probe read a deck back through it on PowerPoint
 * for the web. Each slice is turned into a string in chunks because
 * `String.fromCharCode.apply` on a whole slice overruns the argument limit on a
 * deck of any size.
 */
function readDeckThroughFileApi(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(res.error ? res.error.message : "the host would not open the file"));
        return;
      }
      const file = res.value;
      const chunks: string[] = [];
      const next = (i: number): void => {
        if (i >= file.sliceCount) {
          file.closeAsync();
          resolve(btoa(chunks.join("")));
          return;
        }
        file.getSliceAsync(i, (slice) => {
          if (slice.status !== Office.AsyncResultStatus.Succeeded) {
            file.closeAsync();
            reject(new Error(slice.error ? slice.error.message : `the host refused slice ${i}`));
            return;
          }
          const bytes = slice.value.data as number[];
          let s = "";
          for (let j = 0; j < bytes.length; j += 0x8000)
            s += String.fromCharCode.apply(null, bytes.slice(j, j + 0x8000));
          chunks.push(s);
          next(i + 1);
        });
      };
      next(0);
    });
  });
}

export interface InsertOutcome extends InsertVerdict {
  /** The deck's size before and after, measured in their own batches. */
  before: number;
  after: number;
}

/**
 * Insert a deck after the last slide.
 *
 * `expected` is how many slides the package holds, and it is the caller's job
 * to know: the engine built the package and can count its `sldIdLst`, where
 * this file would have to open a zip to find out.
 */
export async function insertDeck(base64: string, expected: number): Promise<InsertOutcome> {
  const before = await slideCount();
  const targetSlideId = await lastSlideId();
  let error: string | undefined;
  try {
    await withTimeout(
      PowerPoint.run(async (context) => {
        // Without a target the host inserts at the FRONT. A sibling project put
        // 37 generated slides ahead of somebody's title slide that way.
        context.presentation.insertSlidesFromBase64(base64, {
          formatting: "KeepSourceFormatting",
          targetSlideId,
        });
        await context.sync();
      }),
      BUDGET.insert,
      "inserting the merged deck",
    );
  } catch (e) {
    error = e instanceof Error ? e.message : String(e);
  }
  const after = await slideCount();
  return { ...insertVerdict({ before, after, expected, error }), before, after };
}

export interface UndoOutcome {
  removed: number;
  detail: string;
}

/**
 * Take back the slides a run added, by position.
 *
 * `deckAtStart` must be the count taken BEFORE the run inserted anything, and
 * `added` what it believes it added. `sweepPlan` refuses any plan whose first
 * index is not at or after `deckAtStart`, so nothing the user owned before the
 * run can be reached even if both numbers are wrong.
 */
export async function undoInsert(deckAtStart: number, added: number): Promise<UndoOutcome> {
  const deckNow = await slideCount();
  const plan = sweepPlan({ deckAtStart, deckNow, added });
  if (!plan) {
    return { removed: 0, detail: `nothing to take back (deck was ${deckAtStart}, is ${deckNow})` };
  }
  await withTimeout(
    PowerPoint.run(async (context) => {
      // Highest index first: removing a slide shifts every index after it, so
      // walking upward would delete the wrong slides after the first.
      for (let i = plan.from + plan.count - 1; i >= plan.from; i--) {
        context.presentation.slides.getItemAt(i).delete();
      }
      await context.sync();
    }),
    BUDGET.undo,
    "removing the slides this run added",
  );
  // A queued delete that raised nothing has not necessarily happened. Adds,
  // inserts and tag writes have all been accepted here and not performed, so
  // the deck is counted again rather than the call being believed.
  const deckAfter = await slideCount();
  const removed = deckNow - deckAfter;
  return {
    removed,
    detail:
      removed === plan.count
        ? `removed ${removed} slide(s) from index ${plan.from}`
        : `asked for ${plan.count} slide(s) from index ${plan.from} and the deck shrank by ${removed}`,
  };
}
