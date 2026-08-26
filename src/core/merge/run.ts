/**
 * Carry out a plan against a package.
 *
 * Everything happens in the file: clone the template slide, replace the
 * placeholders in the copy, write the merge tags into the copy. Nothing is
 * asked of PowerPoint, which is the point. On the web a slide the run has just
 * added does not resolve by id and a tag write through a shape proxy is
 * refused, so metadata that goes in afterwards is metadata that may not go in
 * at all.
 */
import type { RecordSet } from "../data/recordset.js";
import { cloneSlide, type CloneOptions } from "../pptx/clone.js";
import type { Pkg } from "../pptx/pkg.js";
import { writeSlideTags } from "../pptx/tags.js";
import type { MergePlan } from "./plan.js";
import { makeResolver, type EmptyPolicy } from "./resolve.js";
import { mergeDocument } from "./text.js";

export interface RunOptions {
  onEmpty?: EmptyPolicy;
  /** Passed through to the cloner. Injectable so a test can assert on creation ids. */
  clone?: CloneOptions;
}

export interface RunResult {
  runId: string;
  /** Package paths of the slides produced, in the order they were added. */
  slides: string[];
  /** How many paragraphs were rewritten. A zero here on a real template means the fields never matched. */
  paragraphsMerged: number;
}

/**
 * Run the plan and return what it produced.
 *
 * The order the steps arrive in is the order the slides land in, because every
 * clone is appended to the deck's own slide list. The plan is record-major, so
 * a record's slides stay together.
 */
export async function runPlan(
  pkg: Pkg,
  plan: MergePlan,
  records: RecordSet,
  opts: RunOptions = {},
): Promise<RunResult> {
  const slides: string[] = [];
  let paragraphsMerged = 0;

  for (const step of plan.steps) {
    const row = records.rows[step.recordIndex];
    if (!row) continue;

    // Clone first, then merge the COPY. Merging the template and cloning after
    // would leave the template holding one record's values, which is the
    // template destroyed rather than used.
    const target = await cloneSlide(pkg, step.source, opts.clone ?? {});
    paragraphsMerged += mergeDocument(await pkg.doc(target), makeResolver(row, { onEmpty: opts.onEmpty }));
    await writeSlideTags(pkg, target, step.tags);
    slides.push(target);
  }

  return { runId: plan.runId, slides, paragraphsMerged };
}
