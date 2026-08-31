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
 * The separators a paste can arrive with, in the order they are preferred.
 *
 * **Tab first**, because the commonest input by far is a range copied out of
 * Excel and pasted into the pane, and that arrives tab-separated.
 *
 * **Semicolon last, and it is new.** Excel writes `;` as its CSV separator on
 * any machine whose locale uses the comma as a decimal point — Danish, German,
 * French — and such a paste used to read as ONE column named `Navn;Beløb`,
 * with every placeholder unmatched. The order only breaks ties; which one wins
 * is decided by `chooseDelimiter` below.
 */
const DELIMITERS: readonly string[] = ["\t", ",", ";"];

/**
 * How many rows the sniff looks at.
 *
 * The header alone is not enough — see `chooseDelimiter` — and the whole paste
 * is not needed: it is scored three times, on every keystroke in the box. Ten
 * rows is enough for a count to be consistent or not, and it bounds the work
 * on a 240-row paste to something that does not depend on its length.
 */
const SNIFF_ROWS = 10;

/**
 * Which character is separating the cells.
 *
 * **Not "the first candidate the header contains".** That was the rule for a
 * day and it is wrong on the input the semicolon was added for: a Danish sheet
 * headed `Navn;Beløb, EUR;Dato` holds a comma, so the comma won, and
 * `Ada;1,5;2026-01-01` split into `Ada;1` and `5;2026-01-01`. Decimal commas
 * are everywhere in exactly the data that uses semicolons, so a rule that
 * consults only the header meets one at the first opportunity.
 *
 * The rule instead is the one a CSV sniffer uses, and it reads the BODY:
 *
 * 1. Split the sample on each candidate.
 * 2. Keep the candidates that give the SAME number of cells in every row, and
 *    more than one — a separator that is really a separator does that, and a
 *    character that merely appears in the text does not.
 * 3. Of those, take the one giving the MOST cells. `Navn;Beløb, EUR;Dato` is
 *    three columns on the semicolon and two on the comma, and three is the
 *    answer.
 * 4. Tie-break by the order in `DELIMITERS`, and fall back to the comma when
 *    nothing qualifies — a one-column paste has no separator to find, and every
 *    caller wants one column rather than a refusal.
 *
 * Step 4's fallback is what keeps a one-column header holding a semicolon —
 * `Notes; extra` over two more rows — reading as one column: the semicolon
 * splits the first row and not the others, so it is not consistent and never
 * qualifies. The header-only rule split it, and that would have been a
 * regression shipped as a documented trade.
 *
 * **What it still cannot do**, stated because it is a real limit rather than an
 * oversight: `Navn;Beløb, EUR` over one data row is two consistent columns on
 * BOTH characters, and nothing in the text says which. The tie goes to the
 * comma, which is the commoner file worldwide. That is the same shape as the
 * ambiguous slash date this project refuses to guess at — except that here
 * there is no option to refuse, so the tie is broken rather than dodged.
 *
 * **A RAGGED paste is the other limit, and it is left alone deliberately.** A
 * header of three cells over a data row of two is inconsistent on the tab, so
 * the tab is rejected and the whole table reads as one column — the same
 * visible failure the blank-line case had. It is not fixed the same way,
 * because a blank line says nothing about the delimiter while a short row says
 * something ambiguous: the paste may be tab-separated and ragged, or it may
 * genuinely be one column whose text contains tabs. Scoring on a majority
 * instead of on consistency picks a winner where the data does not, which is
 * the rule this file already refuses to break for `Navn;Beløb, EUR`. It comes
 * back if somebody meets it on a real paste, with that paste as the evidence.
 *
 * Scoring with the real splitter is also what retired the header sampler this
 * replaced. That function existed to read the first row with the parser's own
 * quote rule, and it cost two bugs to get right — a quoted newline inside a
 * header cell, and an inch mark in `Size 6" pipe` leaving it inside a quote for
 * the rest of the paste. Both are now handled by construction, because the
 * thing being scored IS the parse. The tests for them are kept and still pass.
 */
function chooseDelimiter(src: string): string {
  let best: { d: string; cells: number } | undefined;
  for (const d of DELIMITERS) {
    // A BLANK LINE is not evidence about the delimiter, and counting it as a
    // row is how a tab-separated paste with a spacer row in it lost its tabs.
    // An empty line splits to one empty cell on every candidate, so it can
    // never match a header of two or more, and step 2 then rejected the real
    // separator — leaving the comma fallback and ONE column whose NAME held the
    // tabs: `Name\tRegion\tRev`, and a merge button that does nothing. The same
    // shape #162 fixed for the semicolon, from a different trigger.
    //
    // Skipped rather than scored around, because it says nothing either way: a
    // row that is a single empty cell is a line with no characters in it, which
    // no delimiter could have split differently.
    const rows = splitOn(src, d, SNIFF_ROWS).filter((r) => !(r.length === 1 && r[0] === ""));
    const cells = rows[0]?.length ?? 0;
    if (cells < 2 || !rows.every((r) => r.length === cells)) continue;
    if (!best || cells > best.cells) best = { d, cells };
  }
  return best?.d ?? ",";
}

/** The whole parse, given the delimiter. Stops after `limit` rows when asked. */
function splitOn(src: string, d: string, limit?: number): string[][] {
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
      if (limit !== undefined && rows.length >= limit) return rows;
      row = [];
      cell = "";
    } else cell += c ?? "";
  }
  row.push(cell);
  rows.push(row);
  return rows;
}

export function parseDelimited(text: string, delimiter?: string): string[][] {
  const src = text.replace(/\r\n?/g, "\n").replace(/\n$/, "");
  return splitOn(src, delimiter ?? chooseDelimiter(src));
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
