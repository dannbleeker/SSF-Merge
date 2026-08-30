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
import type { EmptyPolicy } from "./resolve.js";
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
  /**
   * What an empty cell means. `"skip"` drops the whole record.
   *
   * `EmptyPolicy`, not a second spelling of it. This union was written out
   * here and in `office/merge.ts`, and three copies of a closed set is how the
   * one that gains a member is missed.
   */
  onEmpty?: EmptyPolicy;
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
 * Whether a record gets this slide.
 *
 * The rule `buildPlan` applies, exported because the PANE has to answer the
 * same question before there is a plan. It states the number above the merge
 * button — "9 slides added after slide 10, leaving 19 in the deck" — and it was
 * computing it as slides-per-record times rows, which knows nothing about
 * conditions. With one conditional slide and four rows it promised nine slides
 * and the plan built eight, so the sentence a user reads to decide whether to
 * press was over by one, and the deck size it predicted was wrong with it.
 *
 * A condition naming a column the data does not have is NOT a refusal: the
 * slide is emitted and the pane reports the problem, which is why `columns` is
 * asked rather than assumed.
 */
export function slideApplies(
  slide: Pick<BlockSlide, "condition">,
  row: Record<string, string>,
  columns: Set<string>,
): boolean {
  if (slide.condition === undefined) return true;
  if (!columns.has(slide.condition)) return true;
  return isTruthy(row[slide.condition]);
}

/**
 * Whether a record produces no slides at all.
 *
 * `onEmpty: "skip"` drops a record when a field on one of its slides has no
 * value — and only the slides this record's own conditions leave IN, which is
 * the whole reason this is a function rather than a line. It was one, and the
 * bug it fixed is recorded at the call site: reading every slide in the block
 * made a customer with no renewal note vanish over a blank cell on a renewal
 * slide their own condition had already removed.
 *
 * Exported because the PANE has to answer the same question before there is a
 * plan. It states the number above the merge button, and a pane that counted
 * this differently would put the wrong number on the thing being pressed —
 * which has happened here once already, with conditions, and is what
 * `plannedSlides` exists to stop happening twice.
 *
 * Takes only what it reads: a slide's `fields` and its `condition`. The pane
 * has both — `slideFields` off the template read, and the conditions the user
 * chose — and no package paths, so the narrow type is what makes it callable
 * there at all.
 */
export function recordIsSkipped(
  slides: Pick<BlockSlide, "fields" | "condition">[],
  row: Record<string, string>,
  columns: Set<string>,
  onEmpty: EmptyPolicy,
): boolean {
  if (onEmpty !== "skip") return false;
  const wanted = slides.filter((slide) => slideApplies(slide, row, columns));
  return hasEmptyField(wanted, row, columns);
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

    // Conditions decide FIRST, because a field on a slide this record is not
    // getting cannot be a reason to drop the record. It was: `onEmpty: "skip"`
    // read the fields of every slide in the block, so a customer with no
    // renewal note vanished from the deck entirely over a blank cell on the
    // renewal slide their own condition had already left out.
    const wanted: BlockSlide[] = [];
    const left: SkippedSlide[] = [];
    for (const slide of order) {
      if (slide.condition !== undefined && !slideApplies(slide, row, columns)) {
        left.push({ recordIndex, seq: slide.seq, condition: slide.condition });
        continue;
      }
      wanted.push(slide);
    }

    // The RULE, asked rather than restated — `plannedSlides` asks the same
    // function, so the button's number and the plan cannot come apart. It
    // splits by condition again, which costs a filter per record and is the
    // price of there being one answer.
    if (recordIsSkipped(order, row, columns, onEmpty)) {
      skippedRecords.push(recordIndex);
      continue;
    }

    // Only now, so a record dropped whole does not also report the slides its
    // conditions left out. It contributed nothing; saying it contributed two
    // absences would be two answers about one record.
    skippedSlides.push(...left);

    for (const slide of wanted) {
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

/**
 * Whether any field the block refers to is blank for this record.
 *
 * Only a field the data has a COLUMN for. `row[field] ?? ""` cannot tell a
 * blank cell from a column that is not there, and they are different things:
 * a blank cell is data, and a field with no column is an author's typo, which
 * `unknownConditions`' sibling `unmatchedFields` already reports and which
 * `docs/MANUAL.md` promises "always stays on the slide, whatever this control
 * says".
 *
 * Without the check one misspelled placeholder drops EVERY row — the merge
 * deleted by a typo — under a sentence saying in the same breath that the
 * placeholder will stay on the slides, about slides no row would produce.
 */
function hasEmptyField(
  slides: Pick<BlockSlide, "fields">[],
  row: Record<string, string>,
  columns: Set<string>,
): boolean {
  for (const slide of slides) {
    for (const field of slide.fields ?? []) {
      if (!columns.has(field)) continue;
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

/**
 * How many slides each record's steps produced, in plan order.
 *
 * The unit a TORN insert has to be read in. `tornInsert` grades a short insert
 * by walking these and stopping where the slides ran out, so what it needs is
 * one number per record and not a total — "719 of 720 slides landed" is true
 * and useless, because every row after the torn one still looks correct.
 *
 * Grouped from the steps rather than derived from the block's size, and that is
 * the whole reason this exists as a function. The pane has a `slidesPerRecord`
 * that answers `to - from + 1`, which is right for a forecast and wrong here: a
 * condition leaves a row shorter than its neighbours, and a uniform count would
 * then report the tear at the wrong row.
 *
 * Lifted out of `runMerge` on 2026-08-30. It is pure arithmetic over the plan
 * and was reachable only by running a whole merge against a fake host, so the
 * one case worth checking — a record whose condition dropped a slide — could
 * not be checked at all.
 */
export function slidesByRecord(steps: Pick<PlanStep, "recordIndex">[]): number[] {
  const out: number[] = [];
  let lastRecord = -1;
  for (const step of steps) {
    // A NEW GROUP on every change, rather than indexing by `recordIndex`. A
    // plan whose records are contiguous is the only plan `buildPlan` makes, and
    // indexing would silently merge two runs of the same index into one row
    // where this reports them as the two the host would have inserted.
    if (step.recordIndex !== lastRecord) {
      out.push(0);
      lastRecord = step.recordIndex;
    }
    out[out.length - 1] = (out[out.length - 1] ?? 0) + 1;
  }
  return out;
}
