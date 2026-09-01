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
import { provenSweep, sweepPlan } from "../host/undo.js";
import { BUDGET, withTimeout } from "../host/timeout.js";
import { readable } from "../host/errors.js";
import { TAG_RUN } from "../core/pptx/tags.js";

/**
 * What the host says it supports, as the pure layer wants it.
 *
 * **Answers false rather than throwing when it cannot ask.** `Office.context`
 * or its `requirements` being absent is not a supported host reporting an
 * absent set — it is a host that cannot answer the question — and every caller
 * here wants the same thing from both: treat it as unsupported and say so.
 *
 * The guard is at the root because this is called before the pane has decided
 * anything. `ready()` uses it to render the "this PowerPoint is too old"
 * message, and `hostEnvironment()` uses it to say which build was refused; a
 * raise from either leaves the user with a blank pane on precisely the host
 * that needed the sentence. That was live for the length of one commit, when
 * `showBuild()` moved an `isSetSupported` call ahead of the floor check.
 *
 * The rest of `hostEnvironment`'s reads have been individually guarded since it
 * was written, with a docstring saying so. This one was the exception.
 */
export const hostSupports: Supports = (version) => {
  try {
    return Office.context.requirements.isSetSupported("PowerPointApi", version);
  } catch {
    return false;
  }
};

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
    // `String` over a value that may be absent produces the STRING "undefined",
    // which is truthy — so `environmentLine`'s `?? "unknown"` never fired and
    // the line read `platform: "undefined"`, which is the exact outcome that
    // fallback's own comment says it exists to prevent.
    ...((p) => (p ? { platform: p } : {}))(
      read(() => (Office.context.platform === undefined ? undefined : String(Office.context.platform))),
    ),
    // WHICH HOST, from the field that names it. `host` was filled from
    // `diagnostics.version`, so the one field answering "which application am I
    // in" carried a build number and the question went unasked.
    ...((h) => (h ? { host: h } : {}))(
      read(() => {
        const named = Office.context.diagnostics?.host;
        return named === undefined ? undefined : String(named);
      }),
    ),
    ...((v) => (v ? { officeVersion: v } : {}))(read(() => Office.context.diagnostics?.version)),
    supports: hostSupports,
  });
}

/**
 * What names the open document, or "" where the host will not say.
 *
 * The undo crumb lives in `localStorage`, which belongs to the add-in's ORIGIN
 * and is shared by every deck opened against it — so the crumb has to record
 * which deck it was written on, and something has to say which deck is open
 * now. `Office.context.document.url` is the only identity available without a
 * round-trip through the file itself.
 *
 * Answers a string rather than throwing, and "" is a real answer the reader
 * treats as "cannot prove" rather than as a match. It lives HERE rather than
 * beside the crumb because only the pane's entry point may touch Office.js: a
 * pane file that asks the host stops being checkable without a PowerPoint, and
 * the pane is where a wrong label is the thing the user presses.
 */
export function documentKey(): string {
  try {
    return String(Office.context.document.url ?? "");
  } catch {
    return "";
  }
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
    // `readable`, not `e.message`: Office echoes an argument back through
    // `debugInfo`, and the argument here is the ENTIRE merged deck as base64.
    // Uncapped, a failed insert put megabytes of the user's own merged rows on
    // screen as the failure sentence, with the explanation at the front and
    // nothing after it readable — and into a bug report, if they pasted it.
    // `errors.ts` was written for exactly this and its two callers were the
    // paths that REJECT; this one is caught and returned, so it never asked.
    error = readable(e);
  }
  const after = await slideCount();
  return { ...insertVerdict({ before, after, expected, error }), before, after };
}

export interface UndoOutcome {
  removed: number;
  /**
   * Slides in the range the sweep DECLINED — it would not claim them as this
   * run's.
   *
   * Not the same as slides it failed to remove, and the pane read them as the
   * same thing: `added - removed` counted a declined slide as still owed, so
   * the card went on offering to "remove the slides this merge added" over
   * slides the sweep had just said were not this merge's, with a live delete
   * button on them. The notice said both halves at once.
   *
   * **It cannot tell "not ours" from "the host would not say".** `runTagsAt`
   * answers `undefined` for a slide whose tag read came back a null object as
   * well as for one carrying no tag, and `provenSweep`'s own docstring says an
   * unanswered read is not evidence that a slide is not ours. So a host that
   * refuses two of six tag reads produces `disowned: 2`, the pane treats those
   * two as settled, and two of this run's slides stay in the deck with the card
   * gone. The alternative — counting them as still owed — is the defect above,
   * a delete button standing over somebody else's slides, which is worse; but
   * the number is a floor on what is settled rather than a fact about
   * ownership, and nothing downstream may read it as one.
   */
  disowned: number;
  detail: string;
  /**
   * Proof was required, and this host said it cannot give any.
   *
   * Set where `hostSupports("1.3")` — the set that carries `Slide.tags` — is
   * false, so the tag read answered nothing and will answer nothing on every
   * later press. Note what that call actually reports: it is false for a host
   * below 1.3 AND for one this add-in could not ASK, because `Office.context`
   * was missing or `isSetSupported` threw. Both mean the same thing here — no
   * proof is available — which is why they share a flag; neither is a claim
   * about the slides.
   *
   * It is a floor on terminality, not the whole of it. A host that HAS the API
   * and does not answer with it looks identical to one that failed a single tag
   * read, and to a delete the host accepted and performed none of — so it is
   * not reported here, and the pane bounds those by counting presses instead.
   * Withdrawing the offer on the first such answer was a defect of its own: it
   * threw away slides the very next press would have removed.
   */
  unprovable?: boolean;
  /**
   * `sweepPlan` refused the SHAPE of the deck, before any host call.
   *
   * The deck grew by more than this run added, or shrank below where it
   * started, so no window can be named — a co-author's slide, or the user's own
   * editing. Nothing was asked of PowerPoint, so this answer is not evidence
   * about it, and the pane's fruitless-press budget must not be spent on it.
   */
  refusedShape?: boolean;
}

/**
 * The run id each slide in a range says it belongs to.
 *
 * One entry per index, in order; `undefined` where the slide carries no run tag
 * or the host would not answer. Every merged slide gets `SSF_MERGE_RUN` written
 * into the package before the insert, so this is a question the file can answer
 * about itself.
 *
 * Positional throughout — `getItemAt`, never `getItem(id)` — because a slide a
 * run has just added does not round-trip through an id on the web, which is the
 * finding the whole undo path is built around.
 *
 * A host below PowerPointApi 1.3 has no `Slide.tags` at all, and one that does
 * may still refuse. Both come back as "nothing answered", which `provenSweep`
 * reads as "no evidence" rather than as "not ours".
 */
async function runTagsAt(from: number, count: number): Promise<(string | undefined)[]> {
  // ASKED ONLY WHERE IT EXISTS. `Slide.tags` is PowerPointApi 1.3 and this
  // add-in's floor is 1.2, deliberately: the floor is read off the calls the
  // add-in must make, and declaring a higher one turns away hosts that would
  // have run it. This call is not one of those — an undo works without it, just
  // with position as its only evidence — so it is gated rather than required,
  // and a 1.2 host takes the same path as a host that refuses.
  if (!hostSupports("1.3")) return [];
  try {
    return await withTimeout(
      PowerPoint.run(async (context) => {
        const asked = [];
        for (let i = from; i < from + count; i++) {
          const tag = context.presentation.slides.getItemAt(i).tags.getItemOrNullObject(TAG_RUN);
          tag.load("value,isNullObject");
          asked.push(tag);
        }
        await context.sync();
        return asked.map((t) => (t.isNullObject ? undefined : (t.value ?? undefined)));
      }),
      BUDGET.read,
      "asking which slides this run made",
    );
  } catch {
    // Not a failure of the undo. The sweep goes on with position alone, which
    // is what it did before this call existed.
    return [];
  }
}

/**
 * Take back the slides a run added, by position — and only the ones it made.
 *
 * `deckAtStart` must be the count taken BEFORE the run inserted anything, and
 * `added` what it believes it added. `sweepPlan` refuses any plan whose first
 * index is not at or after `deckAtStart`, so nothing the user owned before the
 * run can be reached even if both numbers are wrong.
 *
 * That is a bound on the RANGE and it is not identity. Every clamp in
 * `sweepPlan` compares sizes, so a deck the user has edited to the same total —
 * two merged slides deleted, two of their own appended — passes all of them and
 * yields a plan whose last two entries are slides the user made. `provenSweep`
 * asks the slides themselves, through the run tag the package carries.
 */
export async function undoInsert(
  deckAtStart: number,
  added: number,
  runId: string,
  opts: { requireProof?: boolean } = {},
): Promise<UndoOutcome> {
  const deckNow = await slideCount();
  const plan = sweepPlan({ deckAtStart, deckNow, added });
  if (!plan) {
    // The SHAPE was refused — the deck grew past what this run added, or shrank
    // below where it started — and that says nothing about the host. It is
    // marked so the pane does not spend a press from a budget whose whole
    // purpose is telling a host's hiccup from a host's state: a co-author
    // adding a slide is neither.
    return {
      removed: 0,
      disowned: 0,
      refusedShape: true,
      detail: `nothing to take back (deck was ${deckAtStart}, is ${deckNow})`,
    };
  }
  // PROOF is asked of every host, INCLUDING one that cannot give it.
  //
  // This was gated on `hostSupports("1.3")` for one commit, on the reasoning
  // that `Slide.tags` is 1.3, this add-in's floor is 1.2, and demanding proof
  // of a host with no tags leaves the card sitting over slides no press could
  // ever take. The gate is a slide deletion: on every 1.2 host a second press
  // then falls through to the whole positional window, which is exactly the
  // window that can hold a slide the user made between the presses — and it
  // reports the deletion as a clean success. Guaranteed, for a whole host
  // class, rather than the intermittent case the proof was added for.
  //
  // The dead button it was meant to retire is retired where the decision
  // belongs: a press that removed nothing and disowned something withdraws the
  // offer (see `undoRun` and `endPreview`), which works on a host that cannot
  // answer AND on one that can answer and does not.
  const targets = provenSweep(plan, await runTagsAt(plan.from, plan.count), runId, {
    requireProof: opts.requireProof === true,
  });
  if (targets.length === 0) {
    // Whether the answer can change. Asked HERE because this is the only place
    // that knows both halves — that proof was demanded, and that the host has
    // no tags to give. See `UndoOutcome.unprovable`.
    const unprovable = opts.requireProof === true && !hostSupports("1.3");
    // Says what is KNOWN, which is neither of the two things it used to say.
    // "None of them carries this merge's mark" is a claim about the slides, and
    // `runTagsAt` cannot support it: a slide with no tag and a slide the host
    // would not answer for both arrive as `undefined`, which is the limitation
    // `UndoOutcome.disowned` is documented with. Naming the host instead would
    // be the same guess in the other direction. What is true either way is that
    // none of them could be SHOWN to be this merge's, and that is why nothing
    // went.
    return {
      removed: 0,
      disowned: plan.count,
      detail:
        `nothing to take back — none of slides ${plan.from + 1} to ${plan.from + plan.count} ` +
        `could be shown to be this merge's`,
      ...(unprovable ? { unprovable: true } : {}),
    };
  }
  let error: string | undefined;
  try {
    await withTimeout(
      PowerPoint.run(async (context) => {
        // Highest index first: removing a slide shifts every index after it, so
        // walking upward would delete the wrong slides after the first.
        // `provenSweep` returns them in that order.
        for (const i of targets) context.presentation.slides.getItemAt(i).delete();
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
    error = readable(e);
  }
  // A queued delete that raised nothing has not necessarily happened. Adds,
  // inserts and tag writes have all been accepted here and not performed, so
  // the deck is counted again rather than the call being believed.
  const deckAfter = await slideCount();
  const removed = deckNow - deckAfter;
  const note = error === undefined ? "" : ` (the call raised: ${error})`;
  // Against what was ASKED FOR, which is the proven set and not the whole
  // range: a plan of six that proved four is a complete sweep at four.
  const wanted = targets.length;
  const held = plan.count - wanted;
  // COULD NOT BE SHOWN, never "are not". `held` is `plan.count - targets.length`
  // and a slide the host would not answer for arrives in it exactly like a slide
  // carrying no mark — the limitation `UndoOutcome.disowned` is documented with,
  // and this sentence reaches the user verbatim through the pane's "Nothing was
  // removed — ${detail}". It was the last ownership claim left in the codebase.
  const kept =
    held === 0 ? "" : `; ${held} slide(s) in the range could not be shown to be this merge's and were left alone`;
  // The RAIL's numbering, which is the only one the pane speaks — this string
  // reaches the user verbatim, inside "Nothing was removed — …" and "Some of
  // the merge is still there — …". It said "from index 3" for slides the rail
  // calls 4 to 9: a 0-based index, one before the slides actually touched, in a
  // sentence about slides being deleted. The refusal branch twenty lines above
  // already converts, so one function could report in both numberings at once.
  const range = `slides ${plan.from + 1} to ${plan.from + plan.count}`;
  return {
    removed,
    disowned: held,
    detail:
      removed === wanted
        ? `removed ${removed} slide(s) from ${range}${kept}${note}`
        : `asked for ${wanted} slide(s) from ${range} and the deck shrank by ${removed}${kept}${note}`,
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
    // The deck's ids first, paged and positional — the selection's ids mean
    // nothing without them, and a single collection load of both is exactly
    // what office-js#4272 answers short. See `deckSlideIds`.
    //
    // OUTSIDE the budget below, the way `readTemplate` already does it. Every
    // page of that walk carries its own `BUDGET.read`, and it was nested inside
    // one more of the same size — so the outer budget bounded its own
    // sub-budgets and fired whenever the TOTAL crossed 15 seconds, however
    // promptly the host answered each call. On a 600-slide deck at the round
    // trip times this file's own comment cites, "use the slides I've selected"
    // refused every time with "gave up waiting", after 28 calls that had each
    // answered well inside their budget. A merge of 200 rows over a three-slide
    // block leaves exactly that deck. A wall-clock bound on the whole operation
    // would have to be a larger constant than the per-call one; a budget that
    // also bounds its own parts can only fire falsely.
    const deckIds = await deckSlideIds();
    return await withTimeout(
      PowerPoint.run(async (context) => {
        const selected = context.presentation.getSelectedSlides();
        selected.load("items/id");
        await context.sync();
        return blockFromSelection(
          selected.items.map((s) => s.id),
          deckIds,
        );
      }),
      BUDGET.read,
      "reading the selected slides",
    );
  } catch (e) {
    // An OUTCOME, never a raise. The pane awaits this from a click handler, and
    // a rejection there is an unhandled one — the shape of defect an
    // adversarial review already found in this pane once.
    return { ok: false, why: readable(e) };
  }
}

/** Whether to offer the select-slides shortcut at all. See `canSelectSlides`. */
export function canReadSelection(): boolean {
  return canSelectSlides(hostSupports);
}

/** What became of a press of an Insert button. */
export type CursorInsert = { ok: true } | { ok: false; why: string };

/**
 * Put text where the cursor is, on the slide the user is looking at.
 *
 * `setSelectedDataAsync` is a COMMON API — it carries no PowerPointApi
 * requirement set, so `isSetSupported` cannot answer for it and there is
 * nothing to declare in the manifest. Microsoft documents it as supported in
 * PowerPoint on the web, on Windows and on Mac; what it does NOT document is
 * what happens with no insertion point, which is the ordinary state of a pane
 * the user has just clicked into. So it is guarded at runtime, twice: the
 * method may be absent, and the call may come back with a status this add-in
 * has to turn into a sentence rather than a stack trace.
 *
 * A refusal is an OUTCOME, never a raise. The pane awaits this from a click
 * handler, where a rejection is an unhandled one — and the whole point of the
 * control is that there is a clipboard fallback behind it, which only runs if
 * this answers instead of throwing.
 *
 * Not wrapped in `withTimeout`: this is not `PowerPoint.run`, it takes no
 * batch, and its callback is the host's own. A budget here would produce a
 * false refusal on a host that was about to answer, and the fallback it would
 * send the user to is the worse of the two paths.
 */
export async function insertTextAtCursor(text: string): Promise<CursorInsert> {
  const doc = ((): Office.Document | undefined => {
    try {
      return Office.context.document;
    } catch {
      return undefined;
    }
  })();
  if (!doc || typeof doc.setSelectedDataAsync !== "function") {
    return { ok: false, why: "This PowerPoint will not let the pane type onto a slide." };
  }
  return new Promise<CursorInsert>((resolve) => {
    try {
      doc.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve({ ok: true });
          return;
        }
        // The host's own sentence where there is one. It is the only thing that
        // distinguishes "click into a text box first" from "this host cannot
        // do it at all", and the two want different next moves from the user.
        resolve({ ok: false, why: result.error?.message ?? "PowerPoint would not take the text." });
      });
    } catch (e) {
      // Documented to call back rather than throw, and it throws anyway on a
      // host with no document open. Answered, because the caller has a
      // fallback and a raise would take it with it.
      resolve({ ok: false, why: readable(e) });
    }
  });
}
