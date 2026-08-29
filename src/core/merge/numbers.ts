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
import { Pkg } from "../pptx/pkg.js";
import { C_NS, PKG_REL_NS, R_NS, SSML_NS, elements, parseXml, serializeXml } from "../pptx/xml.js";
import { numericValue } from "../data/format.js";
import { FIELD, type Resolve } from "./text.js";

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
}

export function emptyNumberOutcome(): NumberOutcome {
  return { filled: 0, refused: 0 };
}

export function tallyNumbers(into: NumberOutcome, from: NumberOutcome): void {
  into.filled += from.filled;
  into.refused += from.refused;
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
  const out: string[] = [];
  if (c1 === c2) {
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) out.push(`${columnLetters(c1)}${r}`);
    return out;
  }
  if (r1 === r2) {
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
 * Which worksheet part a sheet TITLE belongs to.
 *
 * The title is in `xl/workbook.xml`, which names it beside an `r:id`; that id
 * is a relationship of the workbook part pointing at the worksheet. Two hops,
 * and the reason the first version of this file skipped them — but the title is
 * the only thing tying a chart's formula to a cell, so guessing instead is
 * guessing which numbers the chart plots.
 */
async function sheetPartFor(book: JSZip, title: string): Promise<string | undefined> {
  const bookFile = book.file("xl/workbook.xml");
  const relsFile = book.file("xl/_rels/workbook.xml.rels");
  if (!bookFile || !relsFile) return undefined;
  let sheets: Element[];
  let rels: Element[];
  try {
    sheets = elements(parseXml(await bookFile.async("string")), SSML_NS, "sheet");
    rels = elements(parseXml(await relsFile.async("string")), PKG_REL_NS, "Relationship");
  } catch {
    return undefined;
  }
  const wanted = sheets.find((s) => s.getAttribute("name") === title);
  const rId = wanted?.getAttributeNS(R_NS, "id") ?? wanted?.getAttribute("r:id");
  if (!rId) return undefined;
  const target = rels.find((r) => r.getAttribute("Id") === rId)?.getAttribute("Target");
  if (!target) return undefined;
  // Targets here are relative to `xl/`, and a leading slash means the package
  // root — the same two shapes a relationship anywhere else can take.
  return target.startsWith("/") ? target.slice(1) : `xl/${target.replace(/^\.\//, "")}`;
}

/** The `<c>` element for one address, or undefined. */
function cellAt(sheet: Document, ref: string): Element | undefined {
  return elements(sheet, SSML_NS, "c").find((c) => c.getAttribute("r") === ref);
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
  if (type === "inlineStr" || type === "str") {
    return elements(cell, SSML_NS, "t")
      .map((t) => t.textContent ?? "")
      .join("");
  }
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
): Promise<NumberOutcome> {
  const out = emptyNumberOutcome();
  if (!workbookPath || !pkg.has(workbookPath)) return out;

  let book: JSZip;
  try {
    book = await JSZip.loadAsync(await pkg.bytes(workbookPath));
  } catch {
    // The same judgement `mergeWorkbook` makes: an embedding that is not a
    // readable zip is a thing to step over, not to lose the run on.
    return out;
  }

  const sheetNames = Object.keys(book.files)
    .filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))
    .sort();
  if (sheetNames.length === 0) return out;
  const sstFile = book.file("xl/sharedStrings.xml");
  let shared: string[];
  const sheets = new Map<string, Document>();
  try {
    shared = sharedStrings(sstFile ? parseXml(await sstFile.async("string")) : undefined);
    for (const name of sheetNames) sheets.set(name, parseXml(await book.file(name)!.async("string")));
  } catch {
    return out;
  }

  const chart = await pkg.doc(chartPath);
  let touched = false;

  for (const ref of elements(chart, C_NS, "numRef")) {
    const formula = elements(ref, C_NS, "f")[0]?.textContent ?? "";
    const cells = cellsOfFormula(formula);
    if (!cells) continue;
    // Resolved once per series rather than per point.
    //
    // A workbook with exactly one sheet falls back to it, because a formula
    // whose title does not resolve is far likelier to be a generator writing an
    // odd name than a real second sheet. With more than one and no match, this
    // gives up and the series keeps its cached numbers — the same refusal
    // `cellsOfFormula` makes for a range it cannot read, and for the same
    // reason: there is no safe guess between two sheets.
    const title = sheetOfFormula(formula);
    const named = title === null ? undefined : await sheetPartFor(book, title);
    const sheetPath = named ?? (sheetNames.length === 1 ? sheetNames[0] : undefined);
    const cache = elements(ref, C_NS, "numCache")[0];
    if (!cache) continue;

    for (const pt of elements(cache, C_NS, "pt")) {
      const idx = Number(pt.getAttribute("idx") ?? "-1");
      const address = cells[idx];
      if (address === undefined) continue;

      // The sheet `<c:f>` NAMES, never whichever one happens to hold the
      // address. The first version searched every sheet and took the first hit,
      // on the reasoning that an embedded chart workbook has one sheet — true
      // of every one this had met, and not a fact about the format. Add a sheet
      // in Edit Data and `B2` exists twice; the merge would then write a number
      // into whichever came first and the chart would plot the other. Silently,
      // and no count would catch it: the right NUMBER of cells is written.
      if (!sheetPath) continue;
      const doc = sheets.get(sheetPath);
      const cell = doc ? cellAt(doc, address) : undefined;
      if (!cell) continue;

      const text = stringOfCell(cell, shared);
      if (text === undefined) continue; // already a number: not ours to touch
      FIELD.lastIndex = 0;
      if (!FIELD.test(text)) continue; // a label somebody typed, left alone

      // Resolved WITHOUT a format on purpose. A number written into a chart is
      // formatted by the chart, and `{{Revenue|number:0}}` would hand back
      // "1 250 000" — which is the right string and an unplottable cell.
      FIELD.lastIndex = 0;
      const filled = text.replace(FIELD, (whole, name: string) => resolve(name) ?? whole);
      const value = numericValue(filled);
      if (value === undefined) {
        out.refused++;
        continue;
      }

      writeNumber(doc!, cell, value);
      const v = elements(pt, C_NS, "v")[0];
      if (v) v.textContent = String(value);
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
