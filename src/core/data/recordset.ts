/**
 * The one shape every data source answers with.
 *
 * A pasted range, a parsed .xlsx, a Graph-read table and a JSON endpoint all
 * end here, and the engine cannot tell them apart. Values are kept as strings
 * plus a detected type rather than coerced on the way in: the merged output is
 * text, and a number that has been through a float is a number that has lost
 * its thousands separator and sometimes its last digit.
 */
import { dateShape, numericValue } from "./format.js";

export type ColumnType = "text" | "number" | "date" | "image";

export interface Column {
  name: string;
  type: ColumnType;
}

export interface RecordSet {
  columns: Column[];
  /** One entry per row, keyed by column name. Missing and blank are both "". */
  rows: Record<string, string>[];
}

/**
 * Whether a cell is a number we are willing to claim.
 *
 * Defined AS the parser, not beside it. This pair has disagreed three times —
 * `Number()` accepting `0x10` where the pattern refused it, and a grouping
 * pattern admitting `1,234,5` that `numericValue` could not read — and each
 * time the symptom was a column typed here, converted there, and rendered half
 * formatted and half raw with nothing saying why.
 *
 * A sweep over 6190 arrangements of digits and separators used to assert they
 * agreed. It still runs, and it is now a tautology, which is the right end
 * state for a property that should be structural.
 */
export function looksLikeNumber(value: string): boolean {
  return numericValue(value) !== undefined;
}

/**
 * Whether a cell is a date we are willing to claim.
 *
 * NOT defined as `parseDate`, and the difference is deliberate. A date can be
 * well formed and impossible: `31 Feb 2026` has to pass this and fail the
 * parse, so the author sees what they typed instead of a silently corrected
 * day. `dateShape` is that shape test, and it lives beside `parseDate` because
 * the parser is its only other reader — which is what a private copy of the
 * pattern in `parseDate` failed to be until it typed a column `date` and then
 * rendered it raw.
 */
export function looksLikeDate(value: string): boolean {
  return dateShape(value);
}

/**
 * A cell that names an image FILE.
 *
 * Extension only, and only the four the engine can actually embed — a column of
 * `.svg` or `.heic` names would be offered as images and then fail one row at a
 * time, which is worse than never offering it. The name may carry a path,
 * because a spreadsheet built from a folder listing routinely does.
 *
 * Deliberately NOT a URL. The bytes come from files the user picks, so a column
 * of `https://…` is text as far as this is concerned, and the pane does not
 * offer to fetch anything.
 */
const IMAGE_NAME = /\.(png|jpe?g|gif|bmp)$/i;

export function detectType(values: string[]): ColumnType {
  const filled = values.filter((v) => v.trim() !== "");
  if (!filled.length) return "text";
  if (filled.every(looksLikeNumber)) return "number";
  if (filled.every((v) => looksLikeDate(v))) return "date";
  // Last, because it is the narrowest: a column of file names is not a number
  // and not a date, and asking the other two first costs nothing.
  if (filled.every((v) => IMAGE_NAME.test(v.trim()))) return "image";
  return "text";
}

/** The file names an image column refers to, deduplicated, in first-seen order. */
export function imageNamesIn(rows: Record<string, string>[], column: string): string[] {
  const seen = new Set<string>();
  for (const row of rows) {
    const cell = (row[column] ?? "").trim();
    if (cell !== "") seen.add(cell);
  }
  return [...seen];
}

/**
 * Parse a pasted or uploaded table.
 *
 * Tab first, because the commonest input by far is a range copied out of Excel
 * and pasted into the pane, and that arrives tab-separated. Quoted fields are
 * honoured so a comma inside a company name does not split a row.
 */
/**
 * The first row, as the PARSER will see it.
 *
 * `src.indexOf("\n")` is not where the first row ends. A quoted cell may
 * contain a newline — Excel writes one whenever a cell holds a line break, and
 * it is legal CSV — so the first `\n` in the text can be INSIDE the first
 * header cell. The sniff below then samples half a cell, never reaches the tab
 * that separates the columns, and reads a whole tab-separated table as one
 * column: every placeholder unmatched, the merge button down, and nothing on
 * screen saying why.
 *
 * Walking with the same quote rule the parser uses is the only sample that
 * cannot disagree with it.
 */
function firstRow(src: string): string {
  let quoted = false;
  for (let i = 0; i < src.length; i++) {
    const c = src[i];
    // A bare toggle is enough, and the escaped-quote case needs no branch of
    // its own: `""` toggles twice and nets to zero, which is the same state
    // the parser reaches by skipping both. A branch that cannot change the
    // answer was written here first and removed when a revert proved it could
    // not fail.
    if (c === '"') quoted = !quoted;
    else if (c === "\n" && !quoted) return src.slice(0, i);
  }
  return src;
}

export function parseDelimited(text: string, delimiter?: string): string[][] {
  const src = text.replace(/\r\n?/g, "\n").replace(/\n$/, "");
  const d = delimiter ?? (firstRow(src).includes("\t") ? "\t" : ",");
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;

  for (let i = 0; i < src.length; i++) {
    const c = src[i];
    if (quoted) {
      if (c === '"') {
        if (src[i + 1] === '"') {
          cell += '"';
          i++;
        } else quoted = false;
      } else cell += c;
      continue;
    }
    if (c === '"' && cell === "") quoted = true;
    else if (c === d) {
      row.push(cell);
      cell = "";
    } else if (c === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else cell += c ?? "";
  }
  row.push(cell);
  rows.push(row);
  return rows;
}

/**
 * Build a RecordSet from a parsed table.
 *
 * An unnamed column gets a positional name rather than being dropped, because a
 * template may already bind to it and silently losing a column is the failure
 * this whole layer exists to avoid. Duplicate headers are suffixed for the same
 * reason.
 */
export function toRecordSet(table: string[][], opts: { header?: boolean } = {}): RecordSet {
  const header = opts.header ?? true;
  const first = table[0] ?? [];
  const width = table.reduce((w, r) => Math.max(w, r.length), 0);

  // Dedup against every header the sheet declares, not just the ones already
  // taken. Counting forward alone let a made-up name STEAL one a later column
  // really owns: ["Name", "Name", "Name 2"] produced Name, "Name 2",
  // "Name 2 2", so a template's {{Name 2}} bound to the second "Name" and
  // printed the wrong column on every slide, silently. The same happened with
  // an empty header ahead of a real "Column 1".
  const declared = new Set<string>();
  for (let i = 0; i < width; i++) {
    const raw = (header ? (first[i] ?? "") : "").trim();
    if (raw !== "") declared.add(raw);
  }

  const names: string[] = [];
  for (let i = 0; i < width; i++) {
    const raw = (header ? (first[i] ?? "") : "").trim();
    const invented = raw === "";
    const base = invented ? `Column ${i + 1}` : raw;
    let name = base;
    let n = 2;
    // A candidate is free when nothing earlier has taken it AND it is not a
    // name some other column really declares. Its own header is the one
    // exception: that name IS this column's, so a real "Name" keeps "Name".
    // An INVENTED name has no such claim, which is why the empty-header case
    // must clear `declared` from the very first candidate — otherwise the
    // unnamed first column takes "Column 1" from the real one beside it.
    const owned = (candidate: string) => !invented && candidate === base;
    while (names.includes(name) || (!owned(name) && declared.has(name))) name = `${base} ${n++}`;
    names.push(name);
  }

  const body = header ? table.slice(1) : table;
  const rows = body
    .filter((r) => r.some((c) => c.trim() !== ""))
    .map((r) => {
      const rec: Record<string, string> = Object.create(null);
      names.forEach((name, i) => {
        rec[name] = r[i] ?? "";
      });
      return rec;
    });

  const columns = names.map((name) => ({ name, type: detectType(rows.map((r) => r[name] ?? "")) }));
  return { columns, rows };
}
