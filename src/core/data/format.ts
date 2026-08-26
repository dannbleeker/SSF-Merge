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
  const d = new Date(v);
  if (Number.isNaN(d.getTime())) return undefined;
  // Read back in UTC, then rebuild in UTC.
  //
  // `new Date("1 Mar 2026")` is parsed in the host's LOCAL zone while
  // `formatDate` reads UTC fields, so east of UTC every such date came out a
  // day early: in Europe/Copenhagen — this project's own locale —
  // `1 Mar 2026` rendered as `28 Feb 2026`. CI runs in UTC and the only date
  // assertion used the date-only ISO form, which the spec parses as UTC, so
  // nothing caught it.
  return utcDate(d.getFullYear(), d.getMonth() + 1, d.getDate());
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

export function formatDate(d: Date, pattern: string): string {
  return pattern
    .replace(/yyyy/g, String(d.getUTCFullYear()))
    .replace(/yy/g, pad(d.getUTCFullYear() % 100))
    .replace(/MMM/g, MONTHS_EN[d.getUTCMonth()] ?? "")
    .replace(/MM/g, pad(d.getUTCMonth() + 1))
    .replace(/dd/g, pad(d.getUTCDate()))
    .replace(/\bd\b/g, String(d.getUTCDate()));
}

export function formatNumber(n: number, decimals: number, group: string, point: string): string {
  const fixed = Math.abs(n).toFixed(decimals);
  const [whole = "0", frac] = fixed.split(".");
  const grouped = whole.replace(/\B(?=(\d{3})+(?!\d))/g, group);
  return `${n < 0 ? "-" : ""}${grouped}${frac ? point + frac : ""}`;
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
      const decimals = arg === "" ? 0 : Math.trunc(Number(arg));
      if (!Number.isFinite(decimals) || decimals < 0 || decimals > 100) return raw;
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
