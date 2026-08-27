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
import {
  blockFromSelection,
  blockIds,
  canSelectSlides,
  checkFloor,
  chooseDeckSource,
  environmentLine,
  templateOffset,
  type DeckSource,
  type Environment,
  type SelectedBlock,
  type Supports,
} from "../host/capability.js";
import { insertVerdict, type InsertVerdict } from "../host/verdicts.js";
import { sweepPlan } from "../host/undo.js";
import { BUDGET, withTimeout } from "../host/timeout.js";

/** What the host says it supports, as the pure layer wants it. */
export const hostSupports: Supports = (version) => Office.context.requirements.isSetSupported("PowerPointApi", version);

/** Whether this host can run the add-in at all. Ask before anything else. */
export function ready(): { ok: boolean; detail: string } {
  return checkFloor(hostSupports);
}

/**
 * What this host is, gathered from it.
 *
 * The readings are taken here and the SHAPE is decided in `src/host` — the same
 * split as everything else, so what an environment line contains is checkable
 * without a PowerPoint.
 *
 * Each read is individually guarded. A host that throws reading its own version
 * must not cost the round the rest of the line, and the whole value of this is
 * that it always arrives.
 */
export function hostEnvironment(): Environment {
  const read = (f: () => string | undefined): string | undefined => {
    try {
      return f();
    } catch {
      return undefined;
    }
  };
  return environmentLine({
    // Substituted by Vite at build time. Undefined in the suite and under tsc,
    // which is why `environmentLine` answers "unknown" rather than blank.
    ...(typeof __BUILD_STAMP__ === "string" ? { build: __BUILD_STAMP__ } : {}),
    ...((p) => (p ? { platform: p } : {}))(read(() => String(Office.context.platform))),
    ...((h) => (h ? { host: h } : {}))(read(() => Office.context.diagnostics?.version)),
    supports: hostSupports,
  });
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

/**
 * Slides read per sync when taking the deck's ids.
 *
 * office-js#4272: a collection load of more than ~50 items answers SHORT on the
 * web. A sibling add-in pages every collection read at 20 for that reason, and
 * this is the same number for the same reason.
 */
export const ID_PAGE = 20;

/**
 * Every slide's id, in deck order, without a big collection load.
 *
 * `slides.load("items/id")` is the obvious way and it is the one office-js#4272
 * describes failing: past ~50 items the web host answers with fewer than it
 * has, after a sync that SUCCEEDED. This add-in needs that list twice — to pick
 * a template block's ids, and to turn a selection into slide numbers — and a
 * mail-merge template deck is exactly the kind that gets large.
 *
 * What a short read costs here is worth being precise about, because it is not
 * merely a smaller list. `blockIds` slices by INDEX and `blockFromSelection`
 * calls `indexOf`, so if the ids that come back are not the first n in deck
 * order, both answer the wrong SLIDE NUMBER — silently, and the merge then
 * clones slides nobody chose.
 *
 * So the ids are taken by POSITION instead, in pages: `getItemAt(i)` is a
 * different code path from a collection load and is not subject to its ceiling.
 * `getCount` is a scalar, not a load, and is the authority on how many there
 * are. The probe asks whether a short read is prefix-stable on this host
 * (`deckRead`); this does not wait for the answer, because paging is correct
 * either way and the unpaged read is wrong if the answer is bad.
 */
export async function deckSlideIds(): Promise<string[]> {
  const total = await slideCount();
  const ids: string[] = [];
  for (let start = 0; start < total; start += ID_PAGE) {
    const end = Math.min(start + ID_PAGE, total);
    const page = await withTimeout(
      PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        const handles = [];
        for (let i = start; i < end; i++) {
          const slide = slides.getItemAt(i);
          slide.load("id");
          handles.push(slide);
        }
        await context.sync();
        return handles.map((h) => h.id);
      }),
      BUDGET.read,
      `reading slide ids ${start + 1} to ${end}`,
    );
    ids.push(...page);
  }
  return ids;
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
 * Takes SLIDE NUMBERS — 1-based, the numbering the thumbnail rail shows —
 * because the two routes need different things and only one of them needs ids
 * at all. Deliberately NOT a list of ids: the first version of this signature
 * accepted one, and its only caller built it by counting, so
 * `exportAsBase64Presentation` was handed `["4", "5", "6"]` on a host whose
 * slide ids look like `256#3561048925`. Both sides were `string`, so nothing
 * failed until PowerPoint did. See `blockIds`.
 *
 * The subset route therefore asks the host what the ids ARE, in the same batch
 * it exports with. The whole-deck route needs no ids: the block is still
 * wherever it was, which is what `templateOffset` carries.
 */
export async function readTemplate(block: { from: number; to: number }): Promise<TemplateBytes> {
  const choice = chooseDeckSource(hostSupports);
  const offset = templateOffset(choice.source, block.from - 1);
  if (choice.source === "subset") {
    // Paged and positional, never one big collection load. See `deckSlideIds`.
    const deckIds = await deckSlideIds();
    const base64 = await withTimeout(
      PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        const chosen = blockIds(deckIds, block.from, block.to);
        // Thrown rather than returned: this is a call, and the decision it
        // reports on was made in `src/host` where the suite can check it. The
        // message is already a sentence the pane shows as it stands.
        if (!chosen.ok) throw new Error(chosen.why);
        const bytes = slides.exportAsBase64Presentation(chosen.ids);
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
  let error: string | undefined;
  try {
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
  } catch (e) {
    // Caught, exactly as insertDeck catches its own. A call on this host can
    // raise and still have done the work — the probe's third sheet timed out on
    // an insert whose deck delta showed both slides had landed. Letting the
    // rejection escape skipped the re-count below and told the caller the undo
    // had failed while the user's slides were already gone, with no count of
    // what went. The DELTA is the evidence, never the absence of an error.
    error = e instanceof Error ? e.message : String(e);
  }
  // A queued delete that raised nothing has not necessarily happened. Adds,
  // inserts and tag writes have all been accepted here and not performed, so
  // the deck is counted again rather than the call being believed.
  const deckAfter = await slideCount();
  const removed = deckNow - deckAfter;
  const note = error === undefined ? "" : ` (the call raised: ${error})`;
  return {
    removed,
    detail:
      removed === plan.count
        ? `removed ${removed} slide(s) from index ${plan.from}${note}`
        : `asked for ${plan.count} slide(s) from index ${plan.from} and the deck shrank by ${removed}${note}`,
  };
}

/**
 * The template block the user has SELECTED, in slide numbers.
 *
 * Reads the selection and the deck's own id list in one batch, because
 * `blockFromSelection` needs both: a `SlideRange`'s id is not the deck's id
 * (office-js#2474) and the pane speaks in the numbering the thumbnail rail
 * shows, which is an INDEX into that list.
 *
 * Safe to call on this host, and that is measured rather than assumed. A
 * sibling add-in runs a "selection ladder" — a read, `setSelectedSlides`, a
 * read, `setSelectedShapes([id])`, a read, `setSelectedShapes([])`, a read —
 * in every round of its self-test battery, and **every rung has answered in
 * every archived round**, in 550-710ms, with zero refusals and zero silences.
 * Its "edit the chart the user selected" scenario, which reads
 * `getSelectedSlides` exactly as this does, has never failed one. Both counts
 * stood at 174 of 174 when this was measured on 2026-08-27; the claim is the
 * unbroken run, not the total, because the total moves and nothing here would
 * say so (see `docs/SIBLING.md`). The scenario runs AFTER the ladder, so the
 * read survives even the call that office-js#3698 says wedges the subsystem.
 *
 * This add-in never calls `setSelectedShapes` at all, which is the only call
 * ever implicated in that wedge, so it is on the safest part of a surface that
 * is measured safe.
 */
export async function selectedBlock(): Promise<SelectedBlock> {
  // 1.5, against a floor of 1.2. Asked here as well as before the control is
  // drawn, because `selectedBlock` is exported and a caller that skipped the
  // check would otherwise get a TypeError dressed up as a host refusal.
  if (!canSelectSlides(hostSupports)) {
    return { ok: false, why: "This PowerPoint cannot say which slides are selected — type the two numbers instead." };
  }
  try {
    return await withTimeout(
      (async () => {
        // The deck's ids first, paged and positional — the selection's ids mean
        // nothing without them, and a single collection load of both is exactly
        // what office-js#4272 answers short. See `deckSlideIds`.
        const deckIds = await deckSlideIds();
        return PowerPoint.run(async (context) => {
          const selected = context.presentation.getSelectedSlides();
          selected.load("items/id");
          await context.sync();
          return blockFromSelection(
            selected.items.map((s) => s.id),
            deckIds,
          );
        });
      })(),
      BUDGET.read,
      "reading the selected slides",
    );
  } catch (e) {
    // An OUTCOME, never a raise. The pane awaits this from a click handler, and
    // a rejection there is an unhandled one — the shape of defect an
    // adversarial review already found in this pane once.
    return { ok: false, why: e instanceof Error ? e.message : String(e) };
  }
}

/** Whether to offer the select-slides shortcut at all. See `canSelectSlides`. */
export function canReadSelection(): boolean {
  return canSelectSlides(hostSupports);
}
