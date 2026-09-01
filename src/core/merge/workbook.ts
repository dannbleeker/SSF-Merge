/**
 * Which parts of an EMBEDDED WORKBOOK hold text a merge may fill.
 *
 * A chart's data is a whole `.xlsx` sitting inside the `.pptx`, and two passes
 * read it: `mergeChartNumbers` turns a value cell back into a number, and
 * `mergeWorkbook` fills the placeholders in everything else. They have to agree
 * about which parts those are, and they did not.
 *
 * The numeric pass took the worksheets from what the workbook DECLARES, and its
 * own comment says why: "Excel writes that name and other generators do not,
 * and a workbook whose sheet is called anything else filled NOTHING — not a
 * refusal, not a count, no mention." The text pass matched
 * `xl/worksheets/sheetN.xml` by pattern — the reader that lesson was written
 * about — one file over, and the population it excludes is the very one it
 * reads worksheets FOR: a cell holds its string inline rather than in the
 * shared-string table when the workbook was written by a tool rather than by
 * Excel, and a tool that names its cells differently names its parts
 * differently too.
 *
 * Measured on one file: a workbook whose sheet is `xl/sheets/data.xml`, holding
 * a value cell and a label cell side by side, came back with the value filled
 * and the label still reading `{{Name}}`. Nothing said so — the deck's own
 * chart cache was right, and the placeholder only shows when somebody presses
 * Edit Data.
 *
 * So both sides ask this instead. A shared function is the only version of
 * "these two agree" that cannot rot.
 *
 * Part names are read one hop at a time, all the way down, because every one of
 * them is arbitrary: `_rels/.rels` names the workbook part, the workbook names
 * its sheets, and its relationships name the shared-string table. The
 * conventional spelling is kept as a FALLBACK rather than as the rule, so a
 * workbook whose relationships are missing or unreadable still reads the way it
 * always did.
 */
import type JSZip from "jszip";
import { Pkg, resolveTarget } from "../pptx/pkg.js";
import { PKG_REL_NS, R_NS, SSML_NS, elements, parseXml } from "../pptx/xml.js";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/** The package's own pointer at the workbook part. */
const OFFICE_DOCUMENT = `${REL}/officeDocument`;
/** The shared-string table, where Excel keeps every string a cell shows. */
const SHARED_STRINGS = `${REL}/sharedStrings`;

/** Where the parts Excel writes live, when nothing says otherwise. */
const CONVENTIONAL_MAIN = "xl/workbook.xml";
const CONVENTIONAL_SST = "xl/sharedStrings.xml";

/**
 * A sheet title as `byTitle` keys it.
 *
 * Case-folded, because Excel treats two titles differing only in case as one
 * name: a workbook cannot hold both `Data` and `data`.
 *
 * Private. Nothing outside this file may key that map by hand — `sheetNamed`
 * is the one way to read it, so the fold cannot be applied on one side and
 * forgotten on the other. That is the shape of the defect it replaced.
 */
function foldTitle(title: string): string {
  return title.toLowerCase();
}

/**
 * The part holding the sheet a formula names, or nothing.
 *
 * The ONE reader of `byTitle`, so the way a title is keyed and the way it is
 * looked up cannot come apart. They had: the map was keyed by the declared
 * name and matched exactly, so a chart whose `<c:f>` spelled the title in a
 * different case found no sheet at all — no fill, no refusal, nothing counted,
 * and a workbook with one sheet hid it behind the fallback.
 */
export function sheetNamed(parts: WorkbookParts, title: string): string | undefined {
  return parts.byTitle.get(foldTitle(title));
}

/**
 * How much inflated XML one embedded workbook may cost, in characters.
 *
 * A chart's data sheet is tens of kilobytes; Excel writes nothing near this.
 * The number is three orders of magnitude of headroom, and it is here to bound
 * a DECOMPRESSION BOMB rather than to judge a big workbook.
 *
 * A `.pptx` arrives from wherever the user got it, and a zip entry declares its
 * inflated size before anything inflates it. Measured: 19 KB of deflate becomes
 * 20 MB of text at a ratio of about 1000:1 — and this read happens ONCE PER
 * MERGED ROW, because every clone gets its own copy of the chart's workbook. At
 * 240 rows that is a pane doing gigabytes of work inside a WebView, from a deck
 * somebody was sent.
 *
 * Refusing is the same answer an unparseable workbook already gets: the chart
 * keeps its cached values, the run finishes, and the pane says the data behind
 * it could not be opened. That is a sentence the user can act on, where a
 * frozen tab is not.
 */
export const INFLATED_BUDGET = 64 * 1024 * 1024;

/**
 * Whether a workbook's XML is small enough to read at all.
 *
 * Asked of the zip's own DECLARED sizes, which is the point: it costs nothing
 * and it is answered before a single byte is inflated. Only the XML parts are
 * counted — an embedded image inside a workbook is not something either pass
 * reads, so its size is not this budget's business.
 *
 * `budget` is a parameter so a test can ask the real question — does it sum the
 * declared sizes and refuse past the line — without building a workbook the
 * size of the real budget. The first version of that test allocated eighty
 * megabytes and timed out in CI, which is a slow test rather than a strict one.
 * The constant itself is asserted separately.
 */
export function withinInflatedBudget(book: JSZip, budget = INFLATED_BUDGET): boolean {
  let total = 0;
  for (const name of Object.keys(book.files)) {
    if (!/\.(xml|rels)$/i.test(name)) continue;
    const declared = (book.files[name] as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize;
    if (typeof declared === "number") total += declared;
    if (total > budget) return false;
  }
  return true;
}

export interface WorkbookParts {
  /** Worksheet parts, in the order the workbook declares its sheets. */
  sheets: string[];
  /** Worksheet path by the sheet's own TITLE — what a chart's `<c:f>` names. */
  byTitle: Map<string, string>;
  /** The shared-string table, when the workbook has one. */
  sharedStrings?: string;
}

/** A part's text, or undefined for one that is absent or will not parse. */
async function docOf(book: JSZip, path: string): Promise<Document | undefined> {
  const file = book.file(path);
  if (!file) return undefined;
  try {
    return parseXml(await file.async("string"));
  } catch {
    return undefined;
  }
}

/** `Id` → resolved package path, for one part's relationships. */
async function targetsOf(book: JSZip, ownerPart: string): Promise<Map<string, { path: string; type: string }>> {
  const out = new Map<string, { path: string; type: string }>();
  const doc = await docOf(book, Pkg.relsPathFor(ownerPart));
  if (!doc) return out;
  for (const rel of elements(doc, PKG_REL_NS, "Relationship")) {
    if ((rel.getAttribute("TargetMode") ?? "") === "External") continue;
    const target = rel.getAttribute("Target");
    if (!target) continue;
    out.set(rel.getAttribute("Id") ?? "", {
      // The package's own resolver, never a second reading of what a
      // relationship target means. This had `xl/` glued on by hand once, which
      // is right for the target Excel writes and wrong for the two other
      // shapes one is allowed to take.
      path: resolveTarget(ownerPart, target),
      type: rel.getAttribute("Type") ?? "",
    });
  }
  return out;
}

/**
 * The text-bearing parts of one embedded workbook.
 *
 * Answers empty rather than throwing for anything it cannot read: an embedding
 * under a `package` relationship is not guaranteed to be a workbook at all, and
 * a merge that lost 240 slides over one is worse than a merge that fills
 * nothing in it.
 */
export async function workbookParts(book: JSZip): Promise<WorkbookParts> {
  const root = await targetsOf(book, "");
  const named = [...root.values()].find((r) => r.type === OFFICE_DOCUMENT)?.path;
  const main = named && book.file(named) ? named : CONVENTIONAL_MAIN;

  const rels = await targetsOf(book, main);
  const sheets: string[] = [];
  const byTitle = new Map<string, string>();
  const doc = await docOf(book, main);
  for (const sheet of doc ? elements(doc, SSML_NS, "sheet") : []) {
    const rId = sheet.getAttributeNS(R_NS, "id") ?? sheet.getAttribute("r:id") ?? "";
    const path = rels.get(rId)?.path;
    if (!path) continue;
    const title = sheet.getAttribute("name");
    // Keyed by a FOLDED title, because Excel's sheet names are
    // case-insensitive: a workbook cannot hold both `Data` and `data`, and a
    // chart whose formula spells the title differently from the declaration
    // means the same sheet. A plain `Map` said otherwise, so such a chart's
    // values were never looked at — no fill, no refusal, nothing counted.
    //
    // The single-sheet fallback below hides it entirely, which is why this only
    // shows on a workbook somebody added a sheet to.
    //
    // First declaration wins: two sheets cannot share a title in a workbook
    // Excel will open, and picking the later one for a file that broke that
    // rule would pair a chart's formula with the wrong cells.
    const key = title === null ? null : foldTitle(title);
    if (key !== null && !byTitle.has(key)) byTitle.set(key, path);
    sheets.push(path);
  }

  const related = [...rels.values()].find((r) => r.type === SHARED_STRINGS)?.path;
  const sharedStrings =
    related && book.file(related) ? related : book.file(CONVENTIONAL_SST) ? CONVENTIONAL_SST : undefined;

  return { sheets, byTitle, ...(sharedStrings ? { sharedStrings } : {}) };
}
