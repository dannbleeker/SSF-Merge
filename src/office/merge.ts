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
  let template;
  try {
    template = await readTemplate({ from: req.from, to: req.to });
  } catch (e) {
    // `readTemplate` throws its refusals — `blockIds` produced the sentence and
    // it is already the one to show. A raise here is the host declining to
    // name or export the slides, which is a thing the user can act on.
    return { ok: false, detail: e instanceof Error ? e.message : String(e), fields: [] };
  }
  const pkg = await Pkg.open(template.base64);
  const prepared = await prepareBlock(pkg, { from: req.from, to: req.to, offsetInPackage: template.offset }, "inspect");
  if (!prepared.ok) return { ok: false, detail: prepared.why, fields: [] };
  return {
    ok: true,
    detail: `${prepared.fields.length} placeholder${prepared.fields.length === 1 ? "" : "s"} in slides ${req.from} to ${req.to}.`,
    fields: prepared.fields,
  };
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

  const template = await readTemplate({ from: req.from, to: req.to });
  const pkg = await Pkg.open(template.base64);

  const prepared = await prepareBlock(
    pkg,
    {
      from: req.from,
      to: req.to,
      offsetInPackage: template.offset,
      ...(req.conditions ? { conditions: req.conditions } : {}),
    },
    runId,
  );
  if (!prepared.ok) return { ok: false, detail: prepared.why, ...nothing };

  const plan = buildPlan(prepared.block, req.records, { runId, ...(req.onEmpty ? { onEmpty: req.onEmpty } : {}) });
  if (plan.steps.length === 0) {
    return {
      ok: false,
      detail: "Every row was skipped, so there is nothing to add.",
      ...nothing,
      fields: prepared.fields,
      unknownConditions: plan.unknownConditions,
    };
  }

  const result = await runPlan(pkg, plan, req.records, { ...(req.onEmpty ? { onEmpty: req.onEmpty } : {}) });

  // The package still holds the TEMPLATE slides it cloned from. Inserted, they
  // would put the user's own placeholder slides back into their deck after
  // every run.
  for (const slide of prepared.block.slides) await pkg.removeSlide(slide.path);

  const insert = await insertDeck(await pkg.toBase64(), result.slides.length);
  const added = insert.landed;

  if (insert.verdict !== "yes") {
    return {
      ok: false,
      detail: `The merge was built but PowerPoint did not take it: ${insert.detail}`,
      added,
      deckAtStart,
      runId,
      fields: prepared.fields,
      unknownConditions: plan.unknownConditions,
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
