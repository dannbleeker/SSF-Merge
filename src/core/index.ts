/**
 * The engine's public surface.
 *
 * Nothing here imports Office.js, and a test in `test/architecture.test.ts`
 * holds that. The host layer, the CLI and the task pane are all callers.
 */
export { Pkg, resolveTarget } from "./pptx/pkg.js";
export { cloneSlide, creationIdOf, setCreationId } from "./pptx/clone.js";
export type { CloneOptions } from "./pptx/clone.js";
export {
  TAG_BLOCK,
  TAG_RECORD,
  TAG_RUN,
  TAG_SEQ,
  TAG_TEMPLATE,
  mergeTagPart,
  nextTagNumber,
  readSlideTags,
  tagPartXml,
  writeSlideTags,
} from "./pptx/tags.js";
export { A_NS, CT_NS, PKG_REL_NS, P_NS, R_NS, element, elements, parseXml, serializeXml } from "./pptx/xml.js";
export { FIELD, fieldsIn, mergeDocument, mergeParagraph } from "./merge/text.js";
export type { Resolve } from "./merge/text.js";
export { detectType, looksLikeDate, parseDelimited, toRecordSet } from "./data/recordset.js";
export type { Column, ColumnType, RecordSet } from "./data/recordset.js";
export { applyFormat, formatDate, formatNumber, numericValue, parseDate } from "./data/format.js";
