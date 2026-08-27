/**
 * The merge run: the seam where the pane, the host layer and the engine meet.
 *
 * It orchestrates and decides nothing. Reading the template, inserting and
 * undoing are `powerpoint.ts`; which slides the block is and whether the run may
 * proceed is `prepareBlock`; what the plan is and what it produces is the
 * engine. Every one of those is checked in the suite; this file is the wiring
 * between them and is deliberately small enough to read in one go.
 *
 * The shape it follows is the one the answer sheets earned:
 *
 * 1. count the deck FIRST, because undo is positional and the count taken
 *    before anything was added is the floor that keeps a sweep off the user's
 *    own slides;
 * 2. read the template's bytes;
 * 3. do the whole merge in the package, where nothing can be refused;
 * 4. hand PowerPoint one deck, in one call, anchored after the last slide;
 * 5. read the DELTA to find out what happened, never the absence of an error.
 */
import { runPlan } from "../core/merge/run.js";
import { buildPlan } from "../core/merge/plan.js";
import { prepareBlock } from "../core/merge/prepare.js";
import { Pkg } from "../core/pptx/pkg.js";
import type { RecordSet } from "../core/data/recordset.js";
import type { EmptyPolicy } from "../core/merge/resolve.js";
import { insertDeck, readTemplate, slideCount, undoInsert } from "./powerpoint.js";
import { readable } from "../host/errors.js";
import { trace } from "../core/trace.js";

export interface MergeRequest {
  /** The template block, in the numbering the thumbnail rail shows. */
  from: number;
  to: number;
  records: RecordSet;
  conditions?: Record<number, string>;
  onEmpty?: EmptyPolicy;
  runId?: string;
}

export interface MergeOutcome {
  ok: boolean;
  /** What to show the user, as it stands. */
  detail: string;
  /** Slides added, measured from the deck rather than assumed. */
  added: number;
  /** The deck's size before the insert — what an undo is clamped against. */
  deckAtStart: number;
  runId: string;
  /** Placeholders found in the block, for the pane to report on. */
  fields: string[];
  /** Conditions naming a column the data does not have. */
  unknownConditions: string[];
  /**
   * How many paragraphs the merge actually rewrote.
   *
   * `runPlan` has always returned this, with a docstring saying "a zero here on
   * a real template means the fields never matched", and this file threw it
   * away. So a merge over a template whose placeholders are spelled differently
   * — the single likeliest way a first real run goes wrong — inserted every
   * slide, changed nothing on any of them, and reported unqualified success.
   *
   * The deck delta says how much the deck GREW. This says whether the merge
   * did anything, and they are not the same question.
   */
  paragraphsMerged?: number;
  /** Rows skipped by a condition, so "8 rows" and "6 slides" reconcile. */
  skippedRecords?: number;
  /** Individual slides skipped by a condition. */
  skippedSlides?: number;
}

/** A run id that does not need a clock, so the engine stays deterministic. */
function newRunId(deckAtStart: number, rows: number): string {
  return `r${deckAtStart}-${rows}-${Math.random().toString(36).slice(2, 8)}`;
}

export interface BlockReport {
  ok: boolean;
  /** What to show the user, as it stands. */
  detail: string;
  /** Placeholders found in the block, in the order they appear. */
  fields: string[];
}

/**
 * What a template block holds, without merging anything.
 *
 * The same read and the same preparation `runMerge` does, stopping before the
 * plan. It exists because the pane's fields step is otherwise a guess: the
 * placeholders are in the slides' XML and nothing but reading them says what
 * they are, so a pane that listed nothing would be telling the user their
 * template has no fields when it has six.
 *
 * It costs one template read per press of "Use slides N to M", not one per
 * keystroke — which is why the pane commits the block on the button rather
 * than as the boxes are typed in.
 */
export async function inspectBlock(req: { from: number; to: number }): Promise<BlockReport> {
  // The WHOLE read, not just `readTemplate`. The first version wrapped that one
  // call and awaited `Pkg.open` and `prepareBlock` outside — and both raise: a
  // host that answers the export with bytes JSZip cannot open, or a package
  // with no `ppt/presentation.xml`, which `Pkg.doc` throws for by name. Past
  // the catch those rejections reached `void useBlock()` in the pane with no
  // handler, leaving it saying "Reading the slides…" for the rest of the
  // session with the real error in a console the user cannot open.
  try {
    const template = await readTemplate({ from: req.from, to: req.to });
    const pkg = await Pkg.open(template.base64);
    const prepared = await prepareBlock(
      pkg,
      { from: req.from, to: req.to, offsetInPackage: template.offset },
      "inspect",
    );
    if (!prepared.ok) return { ok: false, detail: prepared.why, fields: [] };
    return {
      ok: true,
      detail: `${prepared.fields.length} placeholder${prepared.fields.length === 1 ? "" : "s"} in slides ${req.from} to ${req.to}.`,
      fields: prepared.fields,
    };
  } catch (e) {
    // `readTemplate` throws its refusals — `blockIds` produced the sentence and
    // it is already the one to show. A raise here is the host declining to
    // name or export the slides, or a package that came back unreadable, and
    // both are things the user can act on.
    return { ok: false, detail: readable(e), fields: [] };
  }
}

export async function runMerge(req: MergeRequest): Promise<MergeOutcome> {
  // Before anything is added. Undo is positional and clamped against this
  // number, so it is taken first and carried through even on the failure
  // paths — a caller that cannot say what the deck was cannot safely take
  // anything back.
  const deckAtStart = await slideCount();
  const runId = req.runId ?? newRunId(deckAtStart, req.records.rows.length);
  const nothing = { added: 0, deckAtStart, runId, fields: [], unknownConditions: [] };

  if (req.records.rows.length === 0) {
    return { ok: false, detail: "There are no rows to merge.", ...nothing };
  }

  // Read and prepare inside a catch, the same way `inspectBlock` does. These
  // were awaited bare, so `readTemplate`'s refusal — a throw this file's own
  // `blockIds` change introduced — escaped `runMerge` entirely and reached the
  // pane's `void merge()` with nothing to catch it. A refusal is an OUTCOME
  // here, like every other one: the caller gets `ok: false` and a sentence, and
  // `deckAtStart` survives so an undo still has its clamps.
  let prepared: Awaited<ReturnType<typeof prepareBlock>>;
  let pkg: Pkg;
  try {
    const template = await readTemplate({ from: req.from, to: req.to });
    pkg = await Pkg.open(template.base64);
    prepared = await prepareBlock(
      pkg,
      {
        from: req.from,
        to: req.to,
        offsetInPackage: template.offset,
        ...(req.conditions ? { conditions: req.conditions } : {}),
      },
      runId,
    );
  } catch (e) {
    return { ok: false, detail: readable(e), ...nothing };
  }
  if (!prepared.ok) return { ok: false, detail: prepared.why, ...nothing };

  const plan = buildPlan(prepared.block, req.records, { runId, ...(req.onEmpty ? { onEmpty: req.onEmpty } : {}) });
  if (plan.steps.length === 0) {
    return {
      ok: false,
      detail: "Every row was skipped, so there is nothing to add.",
      ...nothing,
      fields: prepared.fields,
      unknownConditions: plan.unknownConditions,
      skippedRecords: plan.skippedRecords.length,
      skippedSlides: plan.skippedSlides.length,
    };
  }

  const result = await runPlan(pkg, plan, req.records, { ...(req.onEmpty ? { onEmpty: req.onEmpty } : {}) });

  // Everything that is not a merged slide comes out, and the set is computed
  // from the PACKAGE rather than from the block.
  //
  // Removing `prepared.block.slides` was enough on the `subset` route, where
  // `exportAsBase64Presentation` hands back a package holding only the template
  // block — take the block out and the clones are all that is left. On the
  // `file` route it was not: `getFileAsync` hands back the USER'S ENTIRE
  // PRESENTATION, so the package went to the host as their whole deck, minus
  // the template block, plus the clones. Three rows merged into a forty-slide
  // deck would have inserted forty-six slides, a second copy of everything they
  // had. That route is every host below PowerPointApi 1.10 and this add-in's
  // floor is 1.2, so it is supported and was never exercised.
  //
  // Keeping the clones rather than removing the block makes the two routes one
  // case: on `subset` the difference is exactly the block, so nothing changed
  // there, and `test/office-merge.test.ts` holds both.
  const keep = new Set(result.slides);
  for (const path of await pkg.slidePaths()) {
    if (!keep.has(path)) await pkg.removeSlide(path);
  }

  // What the package HOLDS, not what the plan believed it built. `insertVerdict`
  // grades the deck delta against this number, so taking it from anywhere but
  // the artefact makes the verdict a statement about the wrong thing — and on
  // the file route the two disagreed by the size of the user's deck.
  const sending = await pkg.slidePaths();
  const base64 = await pkg.toBase64();

  // MEASURE THE ARTEFACT BEFORE IT IS GONE.
  //
  // `insertSlidesFromBase64` is handed this string and nothing keeps it: the
  // package goes out of scope the moment the call returns, so when PowerPoint
  // answers `InvalidArgument` the file that caused it no longer exists
  // anywhere. Handing the bytes back is not available either — blob downloads
  // from a task pane do not work (office-js#1511) — so what survives has to be
  // measurements taken while the package is still here.
  //
  // Size especially. A 400-row merge is tens of megabytes of base64 held as a
  // JS string in a task-pane WebView, and it is the number most likely to
  // explain an insert that died inside `BUDGET.insert`. Nothing recorded it.
  //
  // On BOTH populations, not only when the insert fails: a byte count with no
  // successful run to compare it against cannot say whether this one was
  // unusual.
  trace("merge", "handing the package over", {
    slides: sending.length,
    parts: pkg.partNames().length,
    bytes: base64.length,
    rows: req.records.rows.length,
    block: `${req.from}-${req.to}`,
  });

  const insert = await insertDeck(base64, sending.length);
  const added = insert.landed;

  if (insert.verdict !== "yes") {
    // The RECIPE, so the package can be rebuilt without another round.
    //
    // The engine is pure and package-only — `Pkg.open` → `prepareBlock` →
    // `buildPlan` → `runPlan` → `toBase64` reproduces the exact package in
    // Node with no PowerPoint anywhere. So the way to examine a package
    // PowerPoint rejected is not to keep its bytes (which cannot leave a task
    // pane anyway, office-js#1511) but to keep what MADE it. The template is
    // the user's own deck, which they still have.
    //
    // STRUCTURE ONLY, and that is a rule rather than an oversight. A mail
    // merge's rows are the user's confidential data — names, salaries,
    // customer lists — and this record is written to be copied out of the pane
    // and pasted into an issue. Column names and types describe the shape well
    // enough to rebuild a package of the right structure; the values do not
    // change the parts, the relationships or the content types, which is where
    // a rejected package goes wrong.
    trace("merge", "the host refused the package", {
      verdict: insert.verdict,
      block: `${req.from}-${req.to}`,
      rows: req.records.rows.length,
      columns: req.records.columns.map((c) => `${c.name}:${c.type}`),
      conditions: Object.keys(req.conditions ?? {}).length,
      slides: sending.length,
      bytes: base64.length,
      runId,
    });
    return {
      ok: false,
      detail: `The merge was built but PowerPoint did not take it: ${insert.detail}`,
      added,
      deckAtStart,
      runId,
      fields: prepared.fields,
      unknownConditions: plan.unknownConditions,
      paragraphsMerged: result.paragraphsMerged,
      skippedRecords: plan.skippedRecords.length,
      skippedSlides: plan.skippedSlides.length,
    };
  }

  return {
    ok: true,
    detail: `${added} slide${added === 1 ? "" : "s"} added after slide ${deckAtStart}. ${insert.detail}`,
    added,
    deckAtStart,
    runId,
    fields: prepared.fields,
    unknownConditions: plan.unknownConditions,
    // Carried on the SUCCESS path too, not only when something went wrong. A
    // merge that adds every slide and fills no placeholder is the failure this
    // number exists to name, and it arrives looking exactly like a success.
    paragraphsMerged: result.paragraphsMerged,
    skippedRecords: plan.skippedRecords.length,
    skippedSlides: plan.skippedSlides.length,
  };
}

/**
 * Take back what a run added.
 *
 * Positional and clamped, from the count taken before the run started, which is
 * why `MergeOutcome` carries it. Never by id: a slide the run has just added
 * does not round-trip through `slides.getItem(id)` on this host, and a sibling
 * project's by-id clean-up once reported 45 successful deletes having removed
 * nothing.
 */
export async function undoMerge(outcome: MergeOutcome): Promise<{ removed: number; detail: string }> {
  return undoInsert(outcome.deckAtStart, outcome.added);
}
