/**
 * Turn a cell into the string that goes on the slide.
 *
 * Formats are written in the template, next to the field: `{{Amount|number}}`,
 * `{{Start|date:d MMM yyyy}}`, `{{Name|upper}}`. Keeping them in the template
 * rather than in a settings screen means the person who can see the slide is
 * the person who decides how it reads.
 */
import { looksLikeDate } from "./recordset.js";

const MONTHS_EN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
/**
 * For `MMMM`. English on the way OUT whatever the cell was read in, because the
 * pattern is the template author's to choose and they can write the month
 * themselves if they want another language.
 */
const MONTHS_FULL_EN = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

/** Parse the number forms a spreadsheet actually produces, including the European one. */
export function numericValue(raw: string): number | undefined {
  const v = raw.trim();
  if (v === "") return undefined;
  const hasComma = v.includes(",");
  const hasDot = v.includes(".");
  let normalised = v.replace(/\s/g, "");
  if (hasComma && hasDot) {
    // Whichever separator comes last is the decimal one.
    normalised =
      v.lastIndexOf(",") > v.lastIndexOf(".")
        ? normalised.replace(/\./g, "").replace(",", ".")
        : normalised.replace(/,/g, "");
  } else if (hasComma) {
    // "1,5" is a decimal and "1,500" is a thousands group, and nothing in the
    // cell says which. A run of three-digit groups is read as grouping, which
    // is what a spreadsheet exporting a whole number produces; a single comma
    // followed by anything else is the European decimal.
    //
    // `replace` without /g used to change only the FIRST separator, so
    // "1,234,567" became "1234,567" and then NaN — while `detectType` still
    // called the column a number, so half of it rendered formatted and half
    // rendered raw.
    normalised = /^-?\d{1,3}(?:,\d{3})+$/.test(normalised)
      ? normalised.replace(/,/g, "")
      : normalised.replace(",", ".");
  } else if (hasDot && /^-?\d{1,3}(?:\.\d{3})+$/.test(normalised)) {
    // The European grouping spelling, "1.234.567". A single "1.234" stays a
    // decimal: it is genuinely ambiguous and the decimal reading is the one
    // that loses least when wrong.
    normalised = normalised.replace(/\./g, "");
  }
  const n = Number(normalised);
  return Number.isFinite(n) ? n : undefined;
}

export function parseDate(raw: string): Date | undefined {
  if (!looksLikeDate(raw)) return undefined;
  const v = raw.trim();
  const slash = /^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/.exec(v);
  if (slash) {
    const a = Number(slash[1]);
    const b = Number(slash[2]);
    const year = Number(slash[3]);
    // `looksLikeDate` has already refused the ambiguous case, so whichever
    // number cannot be a month is the day.
    const day = a > 12 ? a : b;
    const month = a > 12 ? b : a;
    return utcDate(year < 100 ? 2000 + year : year, month, day);
  }
  // ISO, whose components are right there in the string. Taken from the string
  // rather than from a parsed Date for two reasons, and both have bitten:
  //
  // `new Date("2026-03-01T10:00")` — the form `looksLikeDate` admits — is
  // parsed in the host's LOCAL zone while `formatDate` reads UTC fields, so
  // east of UTC it came out a day early. In Europe/Copenhagen, this project's
  // own locale, `1 Mar 2026` rendered as `28 Feb 2026`.
  //
  // And `new Date` NORMALISES: `2026-02-29` is 1 March, `2026-04-31` is 1 May.
  // Reading the components back off the result cannot catch that, because by
  // then they are the normalised ones and they round-trip perfectly. That is
  // why this function's own `utcDate` guard has always existed and has never
  // fired on this path — it was only ever reached with numbers the parser had
  // already made valid.
  const iso = /^(\d{4})-(\d{2})-(\d{2})(?:[T ]|$)/.exec(v);
  if (iso) return utcDate(Number(iso[1]), Number(iso[2]), Number(iso[3]));

  // `1 Mar 2026` and its Danish and Norwegian spellings. The DAY and the YEAR
  // come from the string; only the month NAME needs the platform, and it is
  // resolved on a day that cannot roll over.
  const named = /^(\d{1,2})[ .\-/]([A-Za-zÆØÅæøå]{3,})[ .\-/](\d{2,4})$/.exec(v);
  if (named) {
    const month = monthFromName(named[2] ?? "");
    if (month === undefined) return undefined;
    const year = Number(named[3]);
    return utcDate(year < 100 ? 2000 + year : year, month, Number(named[1]));
  }
  return undefined;
}

/**
 * Month names in the languages `looksLikeDate` admits, as a stated table.
 *
 * `NAMED_DATE` allows `ÆØÅ`, which is somebody having made room for Danish on
 * purpose. Delegating the name to `new Date` then handled Danish BY ACCIDENT,
 * because that parser matches an English three-letter prefix: `marts` and
 * `januar` worked, `maj` and `desember` did not. One Danish date column came
 * out half formatted and half raw across a merged deck, and the column was
 * still typed `date`, so nothing said why.
 *
 * That is the failure this engine's governing rule is about, in its sharpest
 * form — not a wrong date, but the same column rendering two ways. A table is
 * the smallest thing that makes the answer a decision instead of a side effect
 * of whichever parser the host ships.
 *
 * Full names and the three-letter abbreviations, listed rather than matched by
 * prefix: an open prefix rule reads `1 marketing 2026` as March. No word here
 * means a different month in a different one of these languages, and
 * `test/data.test.ts` asserts that rather than trusting it.
 *
 * This is NOT the Danish locale the backlog rejected. That was a string table
 * for the pane's own text, and it is still rejected. This is reading the user's
 * DATA, which the regex already reached for.
 */
const MONTH_NAMES: Readonly<Record<string, number>> = Object.freeze({
  // English
  january: 1,
  february: 2,
  march: 3,
  april: 4,
  may: 5,
  june: 6,
  july: 7,
  august: 8,
  september: 9,
  october: 10,
  november: 11,
  december: 12,
  // Danish — april, august, september, november and december are the English
  // spellings and are already above.
  januar: 1,
  februar: 2,
  marts: 3,
  maj: 5,
  juni: 6,
  juli: 7,
  oktober: 10,
  // Norwegian differs from Danish in three words.
  mars: 3,
  mai: 5,
  desember: 12,
  // Swedish differs in three more.
  januari: 1,
  februari: 2,
  augusti: 8,
  // The three-letter abbreviations a spreadsheet writes. `okt`, `maj`, `mai`
  // and `des` are the ones an English-only table gets wrong.
  jan: 1,
  feb: 2,
  mar: 3,
  apr: 4,
  jun: 6,
  jul: 7,
  aug: 8,
  sep: 9,
  oct: 10,
  nov: 11,
  dec: 12,
  okt: 10,
  des: 12,
});

/**
 * A month name's number, or nothing if the word is not one we claim.
 *
 * The table first, then `new Date` as a last resort. The fallback is kept
 * deliberately: it is what makes `1 mars 2026` and `1 janvier 2026` work today
 * for languages this table does not list, and removing it would turn a partial
 * answer into no answer for those users. Its inconsistency is exactly why the
 * table exists — it is a floor, not the mechanism.
 *
 * The fallback resolves by parsing the FIRST of that month, which exists in
 * every month of every year, so the answer is the name's month and never a
 * rollover. Asking `new Date` about the whole cell is what let `31 Feb 2026`
 * through as 3 March: the parser had moved the month before anything looked.
 */
function monthFromName(word: string): number | undefined {
  const known = MONTH_NAMES[word.toLowerCase()];
  if (known !== undefined) return known;
  const probe = new Date(`1 ${word} 2001 00:00:00Z`);
  if (Number.isNaN(probe.getTime())) return undefined;
  // The probe is anchored to UTC, so these are the components it was given.
  return probe.getUTCFullYear() === 2001 && probe.getUTCDate() === 1 ? probe.getUTCMonth() + 1 : undefined;
}

/**
 * A UTC date, or nothing if the components are not a real day.
 *
 * `Date.UTC` NORMALISES out of range rather than rejecting: 29 February in a
 * common year becomes 1 March, 31 April becomes 1 May, month 13 becomes January
 * of the next year. Every one of those is a date a reader believes, printed
 * across every merged slide.
 *
 * This engine's governing rule is the opposite, and it is stated in
 * `recordset.ts`: a merged deck that draws perfectly and is two months wrong is
 * worse than one that shows the cell untouched. So the components are read back
 * and the date is refused unless it survived the round trip — which returns the
 * raw cell to the slide, where the author can see it.
 */
function utcDate(year: number, month: number, day: number): Date | undefined {
  const d = new Date(Date.UTC(year, month - 1, day));
  if (Number.isNaN(d.getTime())) return undefined;
  const survived = d.getUTCFullYear() === year && d.getUTCMonth() === month - 1 && d.getUTCDate() === day;
  return survived ? d : undefined;
}

function pad(n: number): string {
  return String(n).padStart(2, "0");
}

/**
 * A date written to a pattern.
 *
 * The replacements run LONGEST TOKEN FIRST, which is what keeps them from
 * eating each other: `yyyy` before `yy`, `MMMM` before `MMM` before `MM`, `dd`
 * before `d`. Reorder them and `yyyy` becomes `2626`.
 *
 * The single `d` is bounded by `\b` so a literal keeps its letters — a pattern
 * of `Ends d` prints "Ends 1" and not "En1s 1". The multi-letter tokens need no
 * such guard: nobody writes `MM` inside a word.
 *
 * Only the tokens the manual lists are tokens. `MMMM` was not one until
 * 2026-08-27 and printed `MarM` — `MMM` replaced and the fourth `M` left
 * standing — which is the shape of every undocumented repeat here.
 */
export function formatDate(d: Date, pattern: string): string {
  return pattern
    .replace(/yyyy/g, String(d.getUTCFullYear()))
    .replace(/yy/g, pad(d.getUTCFullYear() % 100))
    .replace(/MMMM/g, MONTHS_FULL_EN[d.getUTCMonth()] ?? "")
    .replace(/MMM/g, MONTHS_EN[d.getUTCMonth()] ?? "")
    .replace(/MM/g, pad(d.getUTCMonth() + 1))
    .replace(/dd/g, pad(d.getUTCDate()))
    .replace(/\bd\b/g, String(d.getUTCDate()));
}

export function formatNumber(n: number, decimals: number, group: string, point: string): string {
  const fixed = Math.abs(n).toFixed(decimals);
  const [whole = "0", frac] = fixed.split(".");
  const grouped = whole.replace(/\B(?=(\d{3})+(?!\d))/g, group);
  // The sign comes from the ROUNDED value, not the input. `-0.4` at no decimal
  // places rounds to zero, and taking the sign from `n` printed it as `-0` —
  // a quantity that does not exist, on a slide, from an ordinary cell.
  const negative = n < 0 && Number(fixed) !== 0;
  return `${negative ? "-" : ""}${grouped}${frac ? point + frac : ""}`;
}

/**
 * Apply a template's format spec to a raw cell.
 *
 * A value that does not match its format is returned unchanged rather than
 * blanked or replaced with an error marker. The cell is what the user typed,
 * and showing it is more useful on a slide than showing nothing.
 */
export function applyFormat(raw: string, spec: string | undefined): string {
  if (!spec) return raw;
  const [kind = "", arg = ""] = spec.split(":", 2);
  switch (kind.trim().toLowerCase()) {
    case "upper":
      return raw.toLocaleUpperCase();
    case "lower":
      return raw.toLocaleLowerCase();
    case "number": {
      const n = numericValue(raw);
      if (n === undefined) return raw;
      // toFixed throws RangeError outside 0..100, and `number:-1` is a natural
      // thing to write — Excel's ROUND takes negative digits. Thrown, it kills
      // the whole merge with a message naming neither slide nor placeholder,
      // on a path whose own contract is to return the value unchanged.
      // Plain digits only. `Number` also reads `1e2`, `0x10` and ` 2 `, so
      // `number:1e2` asked for a hundred decimal places and got them — a legal
      // `toFixed` call producing a number no slide has room for. A format spec
      // that is not a count of places is not a count of places.
      const trimmed = arg.trim();
      if (trimmed !== "" && !/^\d+$/.test(trimmed)) return raw;
      const decimals = trimmed === "" ? 0 : Number(trimmed);
      if (decimals > 100) return raw;
      return formatNumber(n, decimals, " ", ",");
    }
    case "date": {
      const d = parseDate(raw);
      return d ? formatDate(d, arg || "dd-MM-yyyy") : raw;
    }
    default:
      return raw;
  }
}
