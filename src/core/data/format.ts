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
    // cell says which. Three digits after a single comma is read as a group,
    // which is what a spreadsheet exporting a whole number produces.
    normalised = /^-?\d{1,3},\d{3}$/.test(normalised) ? normalised.replace(",", "") : normalised.replace(",", ".");
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
    const d = new Date(Date.UTC(year < 100 ? 2000 + year : year, month - 1, day));
    return Number.isNaN(d.getTime()) ? undefined : d;
  }
  const d = new Date(v);
  return Number.isNaN(d.getTime()) ? undefined : d;
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
      const decimals = arg === "" ? 0 : Number(arg);
      return formatNumber(n, Number.isFinite(decimals) ? decimals : 0, " ", ",");
    }
    case "date": {
      const d = parseDate(raw);
      return d ? formatDate(d, arg || "dd-MM-yyyy") : raw;
    }
    default:
      return raw;
  }
}
