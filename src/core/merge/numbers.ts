/**
 * A chart's NUMBERS, per recipient.
 *
 * The text merge fills a chart's labels — its title, its series names, its
 * category axis. It cannot fill the values, because a bar's height comes from a
 * `<c:numCache>` cell that has to parse as a number and `{{Revenue}}` does not.
 *
 * The placeholder goes where the value actually lives instead: **the embedded
 * workbook's own cell**, typed in through Edit Data like any other cell. That
 * needs no syntax of its own — the cell becomes an ordinary shared string, and
 * `{{Revenue}}` there means what it means everywhere else.
 *
 * Two things then have to move together, and this pass is the reason they do:
 *
 * - the WORKBOOK cell, from a string cell back to a numeric one, because a
 *   chart plots nothing from text and Excel shows the formula bar to anyone who
 *   presses Edit Data;
 * - the chart's own `<c:numCache>`, because that is the copy PowerPoint draws
 *   from without opening the workbook at all.
 *
 * Filling one and not the other is the half-merge this project already knows
 * about from the text side: the deck looks right until Excel touches it, then
 * reverts in front of the user.
 *
 * `<c:f>` is what joins them. `Sheet1!$B$2:$B$3` beside a cache of two points
 * says point 0 is B2 and point 1 is B3, so the cache index and the cell address
 * are the same fact written twice.
 */
import JSZip from "jszip";
import type { Pkg } from "../pptx/pkg.js";
import { C_NS, CX_NS, SSML_NS, children, elements, parseXml, serializeXml } from "../pptx/xml.js";
import { sheetNamed, withinInflatedBudget, workbookParts } from "./workbook.js";
import { numericValue } from "../data/format.js";
import { fieldsInText, type Resolve } from "./text.js";

export interface NumberOutcome {
  /** Chart values filled from a row. */
  filled: number;
  /**
   * Placeholders in a value cell that did NOT resolve to a number.
   *
   * Counted rather than written, and left exactly as they are. A cell reading
   * `{{Notes}}` cannot become a bar height, and guessing zero would draw a
   * chart the data never said — the same reason a placeholder with no column
   * stays on a slide rather than blanking it.
   */
  refused: number;
  /**
   * Series this pass gave up on before it could look at a cell.
   *
   * Two ways in: the `<c:f>` range could not be read, or it names a sheet the
   * workbook does not declare. Both were a bare `continue` — no fill, no
   * refusal, nothing — so a chart whose values never got as far as being
   * refused reported `{filled: 0, refused: 0}` and the summary said nothing at
   * all. `summary.ts` reasoned from that pair that a zero is always a refusal,
   * and it was not: a third outcome existed and was silent.
   *
   * Per SERIES rather than per cell, because that is the only honest
   * granularity — when the range cannot be read there is no cell to count.
   */
  unreadable: number;
  /**
   * Value placeholders the chart's own cache has no point for.
   *
   * A cached point list is sparse: `<c:ptCount>` covers the range and each
   * `<c:pt>` carries an index, so a writer omits the point for a cell it has no
   * number for — which is exactly the cell somebody typed a placeholder into.
   * The merge fills the data SHEET (an ordinary text merge does that) and has
   * nowhere to put the value in the chart.
   *
   * Counted rather than repaired. Writing the missing point changes what
   * PowerPoint draws, and whether the documented Edit Data route produces a gap
   * at all is unproven. A count turns a silent hole into a sentence with a
   * remedy: the value is in the sheet, and opening Edit Data and closing it
   * brings the chart into line.
   */
  unplotted: number;
}

/**
 * The same counts, plus what the workbook's later text pass must not touch.
 *
 * Separate from `NumberOutcome` because that one is a REPORT — it reaches the
 * pane and is spoken about in sentences — while this is plumbing between two
 * passes of the same merge. Nothing outside `mergeGraphics` has any use for it.
 */
export interface NumberPass extends NumberOutcome {
  /**
   * A refused value cell, named either by the shared string it reads through
   * or — when it carries its own text — by sheet and reference.
   *
   * Counting a refusal does not by itself stop the pass that runs next from
   * merging the very placeholder this one declined. See the refusal itself.
   */
  held: HeldCell[];
}

/** A node the numeric pass has claimed: an `<si>` by index, or a cell by address. */
export type HeldCell = { si: number } | { sheet: string; ref: string };

export function emptyNumberOutcome(): NumberOutcome {
  return { filled: 0, refused: 0, unreadable: 0, unplotted: 0 };
}

/** The same, for the pass that also collects what it declined. */
export function emptyNumberPass(): NumberPass {
  return { filled: 0, refused: 0, unreadable: 0, unplotted: 0, held: [] };
}

export function tallyNumbers(into: NumberOutcome, from: NumberOutcome): void {
  into.filled += from.filled;
  into.refused += from.refused;
  into.unreadable += from.unreadable;
  into.unplotted += from.unplotted;
}

/** The shared-string index a cell reads through, or nothing if it holds its own text. */
function sharedIndexOf(cell: Element): { si: number } | undefined {
  if ((cell.getAttribute("t") ?? "n") !== "s") return undefined;
  const v = elements(cell, SSML_NS, "v")[0]?.textContent ?? "";
  const n = Number(v);
  return Number.isInteger(n) && n >= 0 ? { si: n } : undefined;
}

/** `B` → 2, `AA` → 27. Excel's column letters are base-26 with no zero. */
function columnNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

/** 2 → `B`, 27 → `AA`. */
function columnLetters(n: number): string {
  let out = "";
  let left = n;
  while (left > 0) {
    const rem = (left - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    left = Math.floor((left - 1) / 26);
  }
  return out;
}

/**
 * The cells a chart's `<c:f>` names, in cache-index order.
 *
 * Only a straight line of cells — down one column or across one row — because
 * that is what a chart series is, and because a shape this pass cannot read is
 * a shape it must leave alone rather than guess at. Anything else answers null
 * and the series keeps its cached numbers.
 */
/**
 * The most cells one series may name before this pass refuses to read it.
 *
 * Excel's own ceiling for a 2-D chart is 32,000 points in a single data series,
 * so a range naming more than that is not a series any chart draws — it is a
 * whole-column reference, a generator's mistake, or a hand-written formula.
 *
 * Without a bound the loops below build one string per cell, and the endpoint
 * pattern bounds the SHAPE of an address without bounding its magnitude:
 * `Sheet1!A1:A99999999` allocated until the process died. Not a hang and not an
 * exception the pane could show — a fatal out-of-memory abort, reached from
 * `prepareBlock`, which runs when the user picks the template block and before
 * any data has been pasted.
 *
 * A legal deck reaches the same code: `Sheet1!$A$1:$A$1048576` is what Excel
 * writes when somebody selects a whole column as a series, and it is inside no
 * bound at all. Refusing is the answer this function already gives for a
 * rectangle, and every caller handles it the same way — the series is counted
 * `unreadable` and keeps its cached numbers, which is the conservative outcome.
 */
export const MAX_SERIES_CELLS = 32_000;

export function cellsOfFormula(formula: string): string[] | null {
  const bang = formula.lastIndexOf("!");
  const range = (bang < 0 ? formula : formula.slice(bang + 1)).replace(/\$/g, "");
  const ends = range.split(":");
  const cell = /^([A-Za-z]+)(\d+)$/;
  const from = cell.exec(ends[0] ?? "");
  if (!from) return null;
  if (ends.length === 1) return [`${from[1]!.toUpperCase()}${from[2]}`];
  if (ends.length !== 2) return null;
  const to = cell.exec(ends[1] ?? "");
  if (!to) return null;

  const [c1, r1] = [columnNumber(from[1]!), Number(from[2])];
  const [c2, r2] = [columnNumber(to[1]!), Number(to[2])];
  // FINITE first, and this is not belt and braces — the bound below cannot see
  // an infinity. `Number("999…9")` overflows to Infinity for a long enough run
  // of digits, and `columnNumber` is unbounded base-26 so about 250 letters
  // does the same; `Math.abs(Infinity - Infinity)` is NaN, and `NaN > 32000` is
  // FALSE, so the guard waves it through. What follows then never terminates:
  // `for (let r = Infinity; r <= Infinity; r++)` does not advance, and
  // `columnLetters(Infinity)` appends a character forever.
  //
  // Both were still fatal aborts of the pane after the magnitude bound landed —
  // same signature, same exit code, same route through `prepareBlock` as the
  // defect that bound was written for. A number too large to count exactly is
  // not a cell address, and this answers null for it like any other range it
  // cannot read.
  if (![c1, r1, c2, r2].every((n) => Number.isSafeInteger(n) && n > 0)) return null;
  const out: string[] = [];
  if (c1 === c2) {
    // Refused BEFORE the loop, never by capping it: a truncated range is a
    // different range, and pairing a cache index with the wrong cell is the one
    // outcome this whole pass is written to avoid.
    if (Math.abs(r2 - r1) + 1 > MAX_SERIES_CELLS) return null;
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) out.push(`${columnLetters(c1)}${r}`);
    return out;
  }
  if (r1 === r2) {
    if (Math.abs(c2 - c1) + 1 > MAX_SERIES_CELLS) return null;
    for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) out.push(`${columnLetters(c)}${r1}`);
    return out;
  }
  // A rectangle. A series does not read one, and walking it in the wrong order
  // would pair a value with the wrong point.
  return null;
}

/**
 * The sheet a chart's `<c:f>` names, unquoted, or null.
 *
 * `Sheet1!$B$2:$B$3` and `'My Sheet'!$B$2` both answer their sheet. The last
 * `!` is the separator, because a quoted name may contain one.
 */
export function sheetOfFormula(formula: string): string | null {
  const bang = formula.lastIndexOf("!");
  if (bang < 0) return null;
  const name = formula.slice(0, bang).trim();
  if (name === "") return null;
  // Excel doubles an apostrophe inside a quoted name.
  return name.startsWith("'") && name.endsWith("'") ? name.slice(1, -1).replace(/''/g, "'") : name;
}

/**
 * Every cell in a sheet, by address, built once per document.
 *
 * `cellAt` walked the whole worksheet for EACH address, so reading a series was
 * quadratic in the sheet: a legal `Sheet1!$A$1:$A$1048576` over twenty rows of
 * data took 67 seconds where the same chart with a two-cell range took 194 ms,
 * and a 240-row merge of it extrapolates past thirteen minutes — in a task-pane
 * WebView, with nothing on screen to say why.
 *
 * Keyed on the Document, so it lives exactly as long as the sheet does and
 * cannot outlive a run. Sound only while nothing ADDS a `<c>` to a sheet after
 * a lookup: the one write this pass makes is a `<v>` inside a cell that already
 * exists (see `setCellNumber`), and `test/chart-numbers.test.ts` holds that by
 * scanning the source, because a new cell would be invisible to a stale index
 * and the merge would leave it as the author typed it with nothing said.
 */
const CELL_INDEX = new WeakMap<Document, Map<string, Element>>();

/** The `<c>` element for one address, or undefined. */
export function cellAt(sheet: Document, ref: string): Element | undefined {
  let index = CELL_INDEX.get(sheet);
  if (!index) {
    index = new Map<string, Element>();
    // First wins, matching the `find` this replaces: a malformed sheet naming
    // one address twice reads the same cell it always did.
    for (const c of elements(sheet, SSML_NS, "c")) {
      const at = c.getAttribute("r");
      if (at !== null && !index.has(at)) index.set(at, c);
    }
    CELL_INDEX.set(sheet, index);
  }
  return index.get(ref);
}

/**
 * What a cell SAYS, when what it says is a string.
 *
 * `t="s"` is an index into the shared-string table; `t="inlineStr"` carries the
 * text itself. A cell holding a number answers undefined, which is the ordinary
 * case and the one this pass leaves alone.
 */
function stringOfCell(cell: Element, shared: string[]): string | undefined {
  const type = cell.getAttribute("t") ?? "n";
  if (type === "s") {
    const v = elements(cell, SSML_NS, "v")[0]?.textContent ?? "";
    return shared[Number(v)];
  }
  if (type === "inlineStr") {
    // `<is>` holds the same `<r><t>` runs a shared string does, so a placeholder
    // split by formatting joins here exactly as it does there. Inline is what a
    // generator that never built a shared-string table writes, which is the
    // same population `mergeGraphics` reads worksheets for — a chart written by
    // a tool rather than by Excel.
    return elements(cell, SSML_NS, "t")
      .map((t) => t.textContent ?? "")
      .join("");
  }
  // `t="str"` is NOT read, deliberately, and this used to claim it was.
  //
  // A `str` cell is the cached result of a FORMULA, and it keeps that result in
  // `<v>` rather than in `<t>` — so the line above returned "" for one and the
  // cell was skipped anyway. The claim was false and the behaviour was right,
  // which is the worst pairing: a reader fixing the element name would have
  // made it merge, and `writeNumber` clears every child of the cell, so filling
  // one would delete the `<f>` that produced it. A merge may take a user's
  // placeholder; it may not take their formula.
  //
  // Behaviour is unchanged by saying so — both paths reach the same `continue`
  // one level up — and `test/chart-numbers.test.ts` holds that.
  return undefined;
}

/** Every entry of the shared-string table, joined out of its runs. */
function sharedStrings(doc: Document | undefined): string[] {
  if (!doc) return [];
  return elements(doc, SSML_NS, "si").map((si) =>
    elements(si, SSML_NS, "t")
      .map((t) => t.textContent ?? "")
      .join(""),
  );
}

/**
 * Make a cell hold a number.
 *
 * The `t` attribute goes, because its absence IS the numeric type, and every
 * child goes with it: a cell that kept its `<is>` alongside a new `<v>` is one
 * Excel reads as text with a stray number attached. The style attribute stays —
 * it is the user's formatting, and this pass changes what the cell holds rather
 * than how it looks.
 */
function writeNumber(sheet: Document, cell: Element, value: number): void {
  cell.removeAttribute("t");
  while (cell.firstChild) cell.removeChild(cell.firstChild);
  const v = sheet.createElementNS(SSML_NS, "v");
  v.textContent = String(value);
  cell.appendChild(v);
}

/**
 * One series' cached numbers: the range they came from, and how to rewrite each.
 *
 * A classic chart and a modern one keep the same fact in different markup, and
 * everything between the two — reading the range, finding the cell, resolving
 * the placeholder, refusing what will not parse — is identical. So the shapes
 * are collected into this and the pass below is written once. A second copy of
 * that logic is how the two would come to disagree about which cells hold a
 * placeholder, and `chartValueFields` reports to the pane from the same walk.
 */
export interface CachedSeries {
  /** The `<c:f>` or `<cx:f>` range. Empty when the chart carries literal data
   * and names no cells, which `cellsOfFormula` refuses like any other range it
   * cannot read. */
  formula: string;
  points: { idx: number; write: (value: number) => void }[];
}

/** A classic chart: `<c:numRef>` holding `<c:f>` and a `<c:numCache>` of `<c:pt><c:v>`. */
function classicSeries(chart: Document): CachedSeries[] {
  const out: CachedSeries[] = [];
  for (const ref of elements(chart, C_NS, "numRef")) {
    const cache = elements(ref, C_NS, "numCache")[0];
    if (!cache) continue;
    out.push({
      formula: elements(ref, C_NS, "f")[0]?.textContent ?? "",
      points: elements(cache, C_NS, "pt").map((pt) => ({
        idx: Number(pt.getAttribute("idx") ?? "-1"),
        write: (value) => {
          const v = elements(pt, C_NS, "v")[0];
          if (v) v.textContent = String(value);
        },
      })),
    });
  }
  return out;
}

/**
 * A modern chart: `<cx:numDim>` holding `<cx:f>` and a `<cx:lvl>` of `<cx:pt>`.
 *
 * Three differences from the classic shape, and each one is a way to get this
 * wrong:
 *
 * - a `<cx:pt>` carries its number as TEXT, with no `<c:v>` inside it;
 * - the cache is a `<cx:lvl>`, and a dimension may hold SEVERAL. A multi-level
 *   numeric dimension's range is a rectangle, which `cellsOfFormula` refuses
 *   anyway — but it is refused here too and by name, because "the other refusal
 *   happens to cover it" is the kind of reasoning that stops being true;
 * - the `<cx:f>` is OPTIONAL. A dimension may carry literal data with no
 *   workbook behind it at all, and then there is no cell to fill.
 *
 * Every numeric dimension is taken whatever its `type` — `val`, `size`, `x`,
 * `y` or `colorVal`. A sunburst and a treemap plot from `size` where a waterfall
 * plots from `val`, so a reader that keyed on `val` would fill some modern
 * charts and silently skip others.
 *
 * `children` rather than a descendant search, so each read is tied to the
 * element that owns it: the points of ONE level, the formula of THIS dimension.
 * It is not what keeps a category label safe — a `<cx:strDim>` is not inside a
 * `<cx:numDim>`, so nothing here could reach one either way. What keeps the two
 * apart is starting from `numDim` at all.
 *
 * Exported for the tests that cannot be reached through a deck: nothing this
 * project can author writes a multi-level numeric dimension or a dimension with
 * no formula, and a refusal no test can drive is a refusal nobody can check.
 */
export function modernSeries(chart: Document): CachedSeries[] {
  const out: CachedSeries[] = [];
  for (const dim of elements(chart, CX_NS, "numDim")) {
    const levels = children(dim, CX_NS, "lvl").filter((lvl) => children(lvl, CX_NS, "pt").length > 0);
    if (levels.length !== 1) continue;
    out.push({
      formula: children(dim, CX_NS, "f")[0]?.textContent ?? "",
      points: children(levels[0]!, CX_NS, "pt").map((pt) => ({
        idx: Number(pt.getAttribute("idx") ?? "-1"),
        write: (value) => {
          pt.textContent = String(value);
        },
      })),
    });
  }
  return out;
}

/**
 * Fill one chart's values from the row, in both places that hold them.
 *
 * The chart document is mutated in place, the way every other merge pass here
 * works. The workbook is a package inside the package, so it is opened, changed
 * and written back as bytes — and only written back if something changed, so an
 * untouched chart costs no recompression.
 */
export async function mergeChartNumbers(
  pkg: Pkg,
  chartPath: string,
  workbookPath: string | undefined,
  resolve: Resolve,
): Promise<NumberPass> {
  const out = emptyNumberPass();
  if (!workbookPath || !pkg.has(workbookPath)) return out;

  let book: JSZip;
  try {
    book = await JSZip.loadAsync(await pkg.bytes(workbookPath));
  } catch {
    // The same judgement `mergeWorkbook` makes: an embedding that is not a
    // readable zip is a thing to step over, not to lose the run on.
    return out;
  }
  // And the same budget, for the same reason and at the same point: before a
  // byte is inflated. Both passes open every workbook once per merged row.
  if (!withinInflatedBudget(book)) return out;

  // The workbook's OWN declarations, shared with the text pass — see
  // `workbook.ts`. Reading them by part name is what let one generator's
  // workbook fill nothing at all, silently.
  const workbook = await workbookParts(book);
  if (workbook.sheets.length === 0) return out;
  const sstFile = workbook.sharedStrings ? book.file(workbook.sharedStrings) : undefined;
  let shared: string[];
  const sheets = new Map<string, Document>();
  try {
    shared = sharedStrings(sstFile ? parseXml(await sstFile.async("string")) : undefined);
    for (const name of workbook.sheets) {
      const file = book.file(name);
      if (file) sheets.set(name, parseXml(await file.async("string")));
    }
  } catch {
    return out;
  }

  const chart = await pkg.doc(chartPath);
  let touched = false;

  for (const series of [...classicSeries(chart), ...modernSeries(chart)]) {
    const { formula } = series;
    const cells = cellsOfFormula(formula);
    if (!cells) {
      // Counted, not skipped in silence. See `unreadable`.
      out.unreadable++;
      continue;
    }
    // Resolved once per series rather than per point.
    //
    // A workbook with exactly one sheet falls back to it, because a formula
    // whose title does not resolve is far likelier to be a generator writing an
    // odd name than a real second sheet. With more than one and no match, this
    // gives up and the series keeps its cached numbers — the same refusal
    // `cellsOfFormula` makes for a range it cannot read, and for the same
    // reason: there is no safe guess between two sheets.
    const title = sheetOfFormula(formula);
    // `sheetNamed`, never `byTitle` by hand: the map is keyed by a folded title
    // because Excel's sheet names are case-insensitive, and one reader is what
    // keeps the keying and the lookup from parting company. See `workbook.ts`.
    const named = title === null ? undefined : sheetNamed(workbook, title);
    const sheetPath = named ?? (workbook.sheets.length === 1 ? workbook.sheets[0] : undefined);
    if (!sheetPath) {
      // Counted once for the series, and moved out of the point loop where it
      // used to sit as a bare `continue`. Giving up is the right answer — there
      // is no safe guess between two sheets — but doing it in silence was not:
      // the pane was told nothing and `summary.ts` reasoned that a zero fill
      // with a zero refusal could not happen.
      out.unreadable++;
      continue;
    }

    // WHICH INDEXES THE CACHE HAS. A chart's cached point list is sparse by
    // design — `<c:ptCount>` says how many cells the range covers and `<c:pt>`
    // carries an `idx`, so a writer omits the point for a cell it has no number
    // for. python-pptx and xlsxwriter both do exactly that for a `None`, and a
    // cell somebody typed a placeholder into is not a number.
    //
    // The walk below is over the POINTS, so a cell with no point was never
    // opened: not filled, not refused, not counted, and not reported to the
    // pane as a field. The workbook's text pass fills it all the same, so the
    // data sheet ends up holding the row's figure under a chart that has
    // nowhere to draw it — and on a host that refreshes the cache from the
    // sheet, the bar appears later, out of nowhere.
    //
    // Writing the missing point is the other half of this and is NOT done here:
    // inserting into a cache changes what PowerPoint draws, and whether the
    // documented Edit Data route actually produces a gap is unproven — it is a
    // ten-minute question in a real PowerPoint. What IS done is counting, so a
    // silence becomes a sentence with a remedy in it.
    const plotted = new Set(series.points.map((p) => p.idx));
    for (let idx = 0; idx < cells.length; idx++) {
      if (plotted.has(idx)) continue;
      const address = cells[idx];
      const doc = address ? sheets.get(sheetPath) : undefined;
      const cell = doc ? cellAt(doc, address!) : undefined;
      if (!cell) continue;
      const text = stringOfCell(cell, shared);
      const hits = text === undefined ? [] : fieldsInText(text);
      if (hits.length === 0) continue;
      // NAMED as well as counted, and this is the half that decides whether the
      // block is merged at all. `chartValueFields` is this same walk driven by a
      // recording resolver, and `prepareBlock` refuses a block whose fields come
      // back empty — "every copy would be identical". So a chart whose only
      // placeholder sits in a cell with no cached point reported no field, and
      // the pane told the author to go and type field names onto a slide that
      // already carried one. The documented workflow, refused.
      //
      // The answer is thrown away deliberately: there is no point to write to,
      // which is the whole condition here. The resolver is pure, so asking it
      // costs nothing on a real run and is the recording channel on a dry one.
      for (const hit of hits) resolve(hit.name);
      out.unplotted++;
    }

    for (const point of series.points) {
      const address = cells[point.idx];
      if (address === undefined) continue;

      // The sheet `<c:f>` NAMES, never whichever one happens to hold the
      // address. The first version searched every sheet and took the first hit,
      // on the reasoning that an embedded chart workbook has one sheet — true
      // of every one this had met, and not a fact about the format. Add a sheet
      // in Edit Data and `B2` exists twice; the merge would then write a number
      // into whichever came first and the chart would plot the other. Silently,
      // and no count would catch it: the right NUMBER of cells is written.
      const doc = sheets.get(sheetPath);
      const cell = doc ? cellAt(doc, address) : undefined;
      if (!cell) continue;

      const text = stringOfCell(cell, shared);
      if (text === undefined) continue; // already a number: not ours to touch
      const hits = fieldsInText(text);
      if (!hits.length) continue; // a label somebody typed, left alone

      // Resolved WITHOUT a format on purpose. A number written into a chart is
      // formatted by the chart, and `{{Revenue|number:0}}` would hand back
      // "1 250 000" — which is the right string and an unplottable cell.
      //
      // A field nobody can resolve keeps its braces, exactly as the text pass
      // leaves one visible: the cell then fails `numericValue` and is counted as
      // refused rather than written as a zero.
      let filled = "";
      let read = 0;
      for (const hit of hits) {
        const whole = text.slice(hit.index, hit.index + hit.length);
        filled += text.slice(read, hit.index) + (resolve(hit.name) ?? whole);
        read = hit.index + hit.length;
      }
      filled += text.slice(read);
      const value = numericValue(filled);
      if (value === undefined) {
        out.refused++;
        // HELD, so the workbook's text pass cannot merge what this one refused.
        //
        // This file's own comment said refusing was enough — "counted rather
        // than written, and left exactly as they are" — and it was not. The
        // text pass runs after this one, with the same resolver, over
        // `sharedStrings.xml` and every sheet; a refused cell is still pointing
        // at the entry it merges. So the placeholder the chart kept was
        // rewritten in the sheet: `{{Notes}}` became the row's words, and an
        // empty cell under the default policy became nothing at all.
        //
        // The result is the one thing this pass exists to prevent. The chart
        // goes on drawing the template's number while Edit Data shows something
        // else, and closing Excel refreshes the cache from the sheet — so the
        // bar the user just looked at changes on its own, which the manual
        // promises it does not.
        //
        // What is held is the SHARED STRING where there is one, because that is
        // the node the text pass reaches: skipping the cell alone would leave
        // the `<si>` merged, and one `<si>` may serve several cells. An inline
        // cell holds its own text and is held by reference instead.
        out.held.push(sharedIndexOf(cell) ?? { sheet: sheetPath, ref: address });
        continue;
      }

      writeNumber(doc!, cell, value);
      point.write(value);
      out.filled++;
      touched = true;
    }
  }

  if (!touched) return out;
  for (const [name, doc] of sheets) book.file(name, serializeXml(doc));
  // DEFLATE for the reason `mergeWorkbook` gives: a stored workbook roughly
  // doubles the deck's weight per record, on a package the host swallows whole.
  pkg.setBytes(workbookPath, await book.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
  return out;
}

/**
 * The field names sitting in a chart's VALUE cells.
 *
 * A dry run of `mergeChartNumbers` itself, driven by a resolver that records
 * every name and answers null. Null is what a placeholder with no column gets
 * everywhere else in this engine, and it means the same here: the text is left
 * exactly as it stands, `numericValue` refuses it, nothing is written, and
 * `touched` stays false so no workbook is repacked.
 *
 * The same WALK rather than a second reader, which is the whole point. A
 * scanner that answered this question its own way would be free to disagree
 * with the merge about which cells hold a placeholder — and `prepare.ts` states
 * the rule it would break: "this list and `runPlan`'s are the same list". It is
 * the same list here because it is the same code.
 */
export async function chartValueFields(
  pkg: Pkg,
  chartPath: string,
  workbookPath: string | undefined,
): Promise<string[]> {
  const seen: string[] = [];
  await mergeChartNumbers(pkg, chartPath, workbookPath, (name) => {
    if (!seen.includes(name)) seen.push(name);
    return null;
  });
  return seen;
}
