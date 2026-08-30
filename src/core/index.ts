/**
 * The engine's public surface.
 *
 * Nothing here imports Office.js, and a test in `test/architecture.test.ts`
 * holds that. The host layer, the CLI and the task pane are all callers.
 */
export { Pkg, resolveTarget } from "./pptx/pkg.js";
export { cloneSlide, creationIdOf, notesPathFor, setCreationId } from "./pptx/clone.js";
export type { CloneOptions } from "./pptx/clone.js";
export {
  TAG_BLOCK,
  TAG_RECORD,
  TAG_RUN,
  TAG_SEQ,
  mergeTagPart,
  nextTagNumber,
  readSlideTags,
  tagPartXml,
  writeSlideTags,
} from "./pptx/tags.js";
export {
  A_NS,
  CT_NS,
  PKG_REL_NS,
  P_NS,
  R_NS,
  child,
  children,
  element,
  elements,
  parseXml,
  serializeXml,
} from "./pptx/xml.js";
// `fieldsInText` rather than a pattern: the placeholder syntax is read by a
// linear scan, and handing out a regular expression for callers to iterate is
// what made an unclosed `{{` able to freeze the tab.
export { editRuns, fieldsIn, fieldsInText, mergeDocument, mergeParagraph } from "./merge/text.js";
export type { Edit, FieldHit } from "./merge/text.js";
export { buildPlan, isTruthy, recordCount, slideCount } from "./merge/plan.js";
export type { Block, BlockSlide, MergePlan, PlanOptions, PlanStep, SkippedSlide } from "./merge/plan.js";
export { makeResolver } from "./merge/resolve.js";
export type { EmptyPolicy, ResolveOptions } from "./merge/resolve.js";
export { runPlan } from "./merge/run.js";
export { prepareBlock } from "./merge/prepare.js";
export type { BlockRequest, Prepared } from "./merge/prepare.js";
export type { RunOptions, RunResult } from "./merge/run.js";
export type { Resolve } from "./merge/text.js";
export { detectType, looksLikeDate, parseDelimited, toRecordSet } from "./data/recordset.js";
export type { Column, ColumnType, RecordSet } from "./data/recordset.js";
export { applyFormat, formatDate, formatNumber, numericValue, parseDate } from "./data/format.js";
