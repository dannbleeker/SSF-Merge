/**
 * What to clone, in what order, for which record.
 *
 * A pure decision, deliberately kept away from the package: it takes a block
 * and a set of records and answers a list of steps. That means the ordering
 * rules, the conditional slides and the skipping can be checked without a
 * .pptx anywhere, and the runner that carries them out has no judgement of its
 * own to get wrong.
 */
import type { RecordSet } from "../data/recordset.js";
import { TAG_BLOCK, TAG_RECORD, TAG_RUN, TAG_SEQ } from "../pptx/tags.js";

/** One slide of a template block. */
export interface BlockSlide {
  /** Package path of the template slide, e.g. `ppt/slides/slide4.xml`. */
  path: string;
  /** Position within the block, 1-based, in the deck's own order. */
  seq: number;
  /** Emit this slide only for records where the named field is truthy. */
  condition?: string;
  /** Field names this slide refers to. Scanned by the caller; used by `onEmpty: "skip"`. */
  fields?: string[];
}

/** One or more contiguous slides that repeat together, once per record. */
export interface Block {
  id: string;
  slides: BlockSlide[];
}

export interface PlanStep {
  /** The template slide to clone. */
  source: string;
  blockId: string;
  seq: number;
  /** Index into `records.rows`. */
  recordIndex: number;
  /** Written into the clone before it ever reaches PowerPoint. */
  tags: [string, string][];
}

export interface SkippedSlide {
  recordIndex: number;
  seq: number;
  condition: string;
}

export interface MergePlan {
  runId: string;
  blockId: string;
  steps: PlanStep[];
  /** Records that contributed no slides at all, because `onEmpty` is `"skip"`. */
  skippedRecords: number[];
  /** Individual slides a condition left out. */
  skippedSlides: SkippedSlide[];
  /**
   * Conditions naming a field the data does not have.
   *
   * Reported rather than acted on. A condition nobody can evaluate is an
   * authoring mistake, and dropping the slide would hide it behind output that
   * looks finished; the slide is emitted and the pane says why.
   */
  unknownConditions: string[];
}

export interface PlanOptions {
  /** Which rows to merge. Defaults to all of them, in order. */
  recordIndexes?: number[];
  /** What an empty cell means. `"skip"` drops the whole record. */
  onEmpty?: "blank" | "keep" | "skip";
  /** Injectable so a test can assert on the tags. */
  runId?: string;
}

/**
 * Whether a cell satisfies a slide's condition.
 *
 * Blank is false. So are a few written-out negatives, in English and Danish,
 * because a spreadsheet boolean does not arrive as a boolean: Excel exports it
 * as the word TRUE or FALSE, localised, and a merge that treated "FALSK" as
 * content would emit every slide it was told to leave out. The list is short
 * and stated in the manual rather than guessed at, which is the same rule this
 * engine applies to an ambiguous date.
 */
const FALSE_WORDS = new Set(["false", "falsk", "no", "nej", "off", "0"]);

export function isTruthy(value: string | undefined): boolean {
  const v = (value ?? "").trim();
  if (v === "") return false;
  return !FALSE_WORDS.has(v.toLowerCase());
}

/**
 * Build the plan.
 *
 * Record-major: every slide of record 1, then every slide of record 2. That is
 * what a reader of the finished deck expects, because a record is the thing the
 * deck is about. Slide-major, all the covers followed by all the detail pages,
 * is a different deck and not a default anybody asked for.
 *
 * A conditional slide is skipped IN PLACE. The slides around it keep their
 * order and their sequence numbers, so a record with two slides and a record
 * with three still read the same way.
 */
export function buildPlan(block: Block, records: RecordSet, opts: PlanOptions = {}): MergePlan {
  if (!block.slides.length) throw new Error("ssf-merge: a block must have at least one slide");

  const runId = opts.runId ?? `run-${Math.floor(Math.random() * 0xffff_ffff).toString(36)}`;
  const onEmpty = opts.onEmpty ?? "blank";
  const order = [...block.slides].sort((a, b) => a.seq - b.seq);
  const columns = new Set(records.columns.map((c) => c.name));

  const unknown = new Set<string>();
  for (const slide of order) {
    if (slide.condition !== undefined && !columns.has(slide.condition)) unknown.add(slide.condition);
  }

  const indexes = opts.recordIndexes ?? records.rows.map((_, i) => i);
  const steps: PlanStep[] = [];
  const skippedRecords: number[] = [];
  const skippedSlides: SkippedSlide[] = [];

  for (const recordIndex of indexes) {
    const row = records.rows[recordIndex];
    if (!row) continue;

    if (onEmpty === "skip" && hasEmptyField(order, row)) {
      skippedRecords.push(recordIndex);
      continue;
    }

    for (const slide of order) {
      if (slide.condition !== undefined && columns.has(slide.condition) && !isTruthy(row[slide.condition])) {
        skippedSlides.push({ recordIndex, seq: slide.seq, condition: slide.condition });
        continue;
      }
      steps.push({
        source: slide.path,
        blockId: block.id,
        seq: slide.seq,
        recordIndex,
        tags: [
          [TAG_RUN, runId],
          [TAG_BLOCK, block.id],
          [TAG_SEQ, String(slide.seq)],
          [TAG_RECORD, String(recordIndex)],
        ],
      });
    }
  }

  return { runId, blockId: block.id, steps, skippedRecords, skippedSlides, unknownConditions: [...unknown] };
}

/** Whether any field the block refers to is blank for this record. */
function hasEmptyField(slides: BlockSlide[], row: Record<string, string>): boolean {
  for (const slide of slides) {
    for (const field of slide.fields ?? []) {
      if ((row[field] ?? "").trim() === "") return true;
    }
  }
  return false;
}

/** How many slides the plan will produce. The number the pane puts above the button. */
export function slideCount(plan: MergePlan): number {
  return plan.steps.length;
}

/** How many records contributed at least one slide. */
export function recordCount(plan: MergePlan): number {
  return new Set(plan.steps.map((s) => s.recordIndex)).size;
}
