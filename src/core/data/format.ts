/**
 * Turn a cell into the string that goes on the slide.
 *
 * Formats are written in the template, next to the field: `{{Amount|number}}`,
 * `{{Start|date:d MMM yyyy}}`, `{{Name|upper}}`. Keeping them in the template
 * rather than in a settings screen means the person who can see the slide is
 * the person who decides how it reads.
 */

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

/**
 * A European "1.234,5" and an American "1,234.5" both mean the same thing, and only one of them parses.
 *
 * The grouping separator is captured and the rest of the groups must repeat
 * THAT one, and a decimal part may not reuse it. Without those two conditions
 * the pattern admitted `1,234,5`, which no locale writes and `numericValue`
 * cannot read: the column typed as a number and then rendered raw, which is the
 * disagreement this whole pair of functions exists to prevent.
 *
 * **A space group is four characters, not one.** The locales that group with a
 * space — Swedish, Norwegian, Finnish, French, Polish, Czech, Russian — do not
 * write U+0020. Excel uses the NO-BREAK space, and modern builds use the
 * NARROW no-break space in French, exactly so a number cannot break across a
 * line; a thin space turns up in exports from elsewhere. Copying a formatted
 * cell puts its DISPLAYED text on the clipboard, so what a paste carries is the
 * space Excel chose. Admitting only the ASCII one typed those columns as text,
 * on a cell that looks like a number and in the same population the semicolon
 * delimiter was added for.
 *
 * They are alternatives inside ONE capture, so the backreference still holds:
 * whichever space opens the number, every later group repeats that one.
 * `1<NBSP>234<NNBSP>567` is not a locale, it is a paste that went through
 * something, and it is still refused.
 *
 * A space is never the DECIMAL separator anywhere, which is why the tail stays
 * `[.,]` — `1<NBSP>5` is not one and a half.
 */
const GROUP_SPACE = " \u00a0\u202f\u2009";
// Built from `GROUP_SPACE` so the four spellings of a space group are named
// ONCE. Written inline, the class would have to be repeated wherever the same
// question is asked, which is how the shape gate and the parser drifted apart
// before.
const NUMBER = new RegExp(
  `^-?\\d{1,3}([${GROUP_SPACE}.,])\\d{3}(?:\\1\\d{3})*(?:(?!\\1)[.,]\\d+)?$|^-?\\d+(?:[.,]\\d+)?$`,
);

/**
 * Whether a cell is a number we are willing to claim — the ONE answer.
 *
 * `looksLikeDate` has been the single definition of a date since `format.ts`
 * started importing it. A number had two: this pattern, which `detectType`
 * asks, and `Number()` at the end of `numericValue`, which is far broader. They
 * disagree, and the disagreement is not theoretical — the comment inside
 * `numericValue` records a merge where "half of it rendered formatted and half
 * rendered raw" because a column was typed here and converted there.
 *
 * What `Number()` admits and a spreadsheet does not: `0x10` is sixteen, `0b11`
 * is three, `0o17` is fifteen, `1e3` is a thousand, `+7` and `.5` parse. A cell
 * reading `0x10` is a product code, and turning it into 16 in somebody's deck
 * is the kind of wrong that looks deliberate.
 *
 * Scientific notation is refused with them, which is the one form worth
 * noticing: it means `1.23E+15` stays text. That is not a change — `detectType`
 * has always called such a column text, so the pane has never offered it as a
 * number. This only stops the formatter and the chart writer from disagreeing
 * with that.
 */
function isNumberShape(value: string): boolean {
  return NUMBER.test(value.trim());
}
/** Deliberately narrow. See `looksLikeDate`. */
const ISO_DATE = /^\d{4}-\d{2}-\d{2}(?:[T ].*)?$/;
/**
 * The letters a month NAME may be written with.
 *
 * Latin only, and deliberately not `\p{L}`. This is a SHAPE gate: whatever it
 * admits, `detectType` types as a date, and only then does `monthFromName`
 * decide. So every script it opens is a script whose month names it cannot
 * read, and a column typed `date` and rendered raw is the pathology this file
 * exists to prevent. Latin is the alphabet the table is written in.
 *
 * It used to be `[A-Za-zÆØÅæøå]` — somebody having made room for Danish
 * and stopped there. Nothing Danish is accented, so what that really excluded
 * was `février`, `août`, `décembre`, `März` and `março`: French, German and
 * Portuguese read nine months of twelve, in one column, half formatted and
 * half raw.
 *
 * The ranges are Latin-1 Supplement and Latin Extended-A and -B, with the two
 * MATHEMATICAL signs that sit inside them cut out — U+00D7 and U+00F7 are not
 * letters, and admitting them would let `1 ××× 2026` be typed a date.
 *
 * U+0300-U+036F are the combining marks, so a name arrives in either
 * normalisation: `é` as one code point, or as `e` followed by U+0301, which
 * is what a CSV exported on macOS carries. `monthFromName` normalises before
 * the lookup, and the two halves have to agree — a class that admits the mark
 * without a lookup that folds it types the column and then refuses it.
 */
const MONTH_LETTER = "A-Za-z\\u00C0-\\u00D6\\u00D8-\\u00F6\\u00F8-\\u024F\\u0300-\\u036F";

/**
 * `1 March 2026`, and `1. marts 2026`.
 *
 * The separator may be followed by SPACE, which is not a detail: Danish writes
 * the day as an ordinal, so `1. marts 2026` — the ordinary long form, and the
 * one this project's own owner writes — is a period AND a space. Requiring
 * exactly one separator character admitted `1 marts 2026` and `1.marts 2026`
 * and refused the form anybody actually types.
 *
 * The month-name table was added so Danish dates would be read. This is the
 * shape most of them arrive in, so until now that table was reachable mainly by
 * spellings nobody uses.
 *
 * Nothing is loosened about AMBIGUITY, which is what the refusals below are
 * for. A month spelled out is unambiguous however it is punctuated; `03/01/2026`
 * is not, and is still refused.
 *
 * Private, and in the same file as the only two things that read it —
 * `dateShape` and `parseDate`. It was exported to `format.ts` because the two
 * lived apart, which is how a column got typed `date` by one and refused by the
 * other; now the shape and the parse are one module and there is nothing to
 * export.
 */
// The combining marks in `MONTH_LETTER` are there ON PURPOSE, so the rule
// below is off for this line. It warns that a class holding them can match
// half a grapheme, and half a grapheme is exactly the unit this pattern works
// in: a decomposed `é` is `e` and U+0301, two members of the class, and `{3,}`
// counts both. Dropping them refuses every accented name a macOS export
// writes; matching graphemes instead would mean re-modelling the whole word
// for no gain, since `monthFromName` normalises before the lookup.
// eslint-disable-next-line no-misleading-character-class
const NAMED_DATE = new RegExp(`^(\\d{1,2})[ .\\-/]\\s*([${MONTH_LETTER}]{3,})[ .\\-/]\\s*(\\d{2,4})$`);

/**
 * Whether a cell is a date we are willing to claim.
 *
 * `03/01/2026` is 3 January in Copenhagen and 1 March in Chicago, and nothing
 * in the cell says which. A slash date whose first two numbers could both be a
 * month is refused rather than guessed: a merged deck that draws perfectly and
 * is two months wrong is worse than one that shows the cell untouched. An
 * unambiguous slash date still parses.
 */
export function dateShape(value: string): boolean {
  const v = value.trim();
  if (ISO_DATE.test(v) || NAMED_DATE.test(v)) return true;
  const slash = /^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/.exec(v);
  if (!slash) return false;
  const a = Number(slash[1]);
  const b = Number(slash[2]);
  return !(a <= 12 && b <= 12 && a !== b);
}

/** Parse the number forms a spreadsheet actually produces, including the European one. */
export function numericValue(raw: string): number | undefined {
  const v = raw.trim();
  if (v === "") return undefined;
  // The same question `detectType` asks, asked once. Without this the two
  // disagreed: a column of `0x10` was typed text and formatted 16.
  if (!isNumberShape(v)) return undefined;
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
    // The European grouping spelling, "1.234.567" — and "1.234" with it, which
    // this comment used to claim stayed a decimal. It never did: `+` is one or
    // more, so a single group has always been read as grouping.
    //
    // Which is the right answer, and the one the comma branch above gives for
    // "1,500". A lone three-digit group after a separator is what a spreadsheet
    // exporting a whole number produces, and reading the two separators by
    // opposite rules would be the surprise. It is genuinely ambiguous either
    // way; what matters is that both are read the same and that the answer is
    // written down.
    normalised = normalised.replace(/\./g, "");
  }
  const n = Number(normalised);
  return Number.isFinite(n) ? n : undefined;
}

export function parseDate(raw: string): Date | undefined {
  if (!dateShape(raw)) return undefined;
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
  // The SAME pattern `looksLikeDate` admitted it by. A private copy here is
  // what let the two drift apart: one said "date", the other said no, and the
  // column rendered half formatted and half raw with nothing saying why.
  const named = NAMED_DATE.exec(v);
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
 * `NAMED_DATE` allowed `ÆØÅ` and nothing else beyond ASCII, which is somebody
 * having made room for Danish on purpose and stopping there; `MONTH_LETTER`
 * carries the rest of Latin now, and the accented names are below.
 * Delegating the name to `new Date` handled Danish BY ACCIDENT,
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
 * prefix: an open prefix rule reads `1 marketing 2026` as March. That was not a
 * hypothetical — the `new Date` fallback this table sat in front of was exactly
 * such a rule and did exactly that, for a year, and `monthFromName` records
 * what it cost and why it is gone.
 *
 * No word here means a different month in a different one of these languages,
 * and `test/data.test.ts` asserts that over the whole table rather than
 * trusting it. That property is what lets one table serve eleven languages, and
 * it is the thing to re-check before adding a twelfth.
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
  // German — januar, februar, april, mai, juni, juli, august, september,
  // oktober and november are spellings already above. `M\u00e4rz` is the only
  // German month that is not, and it was the only one this table could not read.
  "m\u00e4rz": 3,
  dezember: 12,
  // Dutch — april, juni, juli, september, oktober, november and december are
  // already above, and januari and februari came in with Swedish.
  maart: 3,
  mei: 5,
  augustus: 8,
  // French. `mars` and `mai` are already above, under Norwegian. The three
  // accented ones are the whole reason `MONTH_LETTER` is wider than ASCII.
  janvier: 1,
  "f\u00e9vrier": 2,
  avril: 4,
  juin: 6,
  juillet: 7,
  "ao\u00fbt": 8,
  septembre: 9,
  octobre: 10,
  novembre: 11,
  "d\u00e9cembre": 12,
  // Spanish
  enero: 1,
  febrero: 2,
  marzo: 3,
  abril: 4,
  mayo: 5,
  junio: 6,
  julio: 7,
  agosto: 8,
  septiembre: 9,
  octubre: 10,
  noviembre: 11,
  diciembre: 12,
  // Italian — marzo, aprile's cousin `abril` and agosto are shared with Spanish.
  gennaio: 1,
  febbraio: 2,
  aprile: 4,
  maggio: 5,
  giugno: 6,
  luglio: 7,
  settembre: 9,
  ottobre: 10,
  dicembre: 12,
  // Portuguese
  janeiro: 1,
  fevereiro: 2,
  "mar\u00e7o": 3,
  maio: 5,
  junho: 6,
  julho: 7,
  setembro: 9,
  outubro: 10,
  novembro: 11,
  dezembro: 12,
});

/**
 * A month name's number, or nothing if the word is not one we claim.
 *
 * The TABLE, and nothing else. There was a `new Date` fallback behind it, kept
 * on purpose so that languages the table did not list still worked; measured on
 * 2026-08-31, it does the one thing the table was written to prevent.
 *
 * `new Date("1 <word> 2001 00:00:00Z")` matches an English three-letter PREFIX,
 * so it answers **March for `marketing`**, January for `janitor`, November for
 * `novel` and December for `decision`. The table's own docstring names that
 * case as the reason its names are listed rather than matched by prefix — and
 * the branch underneath it was an open prefix rule the whole time. A cell
 * reading `1 marketing 2026` was typed `date` and printed as `01-03-2026`, on
 * every merged slide, with nothing anywhere saying a month had been invented.
 *
 * This engine's governing rule decides it: a merged deck that draws perfectly
 * and is two months wrong is worse than one that shows the cell untouched. So
 * the fallback is gone and the answer is a decision rather than a side effect
 * of whichever parser the host ships — which is what the table was for, and it
 * also means the pane and the CLI cannot answer differently.
 *
 * Nothing is lost by it. What the fallback really covered was French and
 * Spanish and Italian, and those are listed above now — more of them than the
 * prefix rule ever reached, since it read `janvier` and `marzo` by luck and
 * answered nothing at all for `enero`, `maart`, `mei` or `gennaio`.
 *
 * The accented spellings are listed now and `MONTH_LETTER` admits them, which
 * closes the one limit this docstring used to state. `NAMED_DATE` allowed
 * `[A-Za-zÆØÅæøå]`, so `février`, `août`, `décembre`, `März` and `março`
 * never reached this function at all and French, German and Portuguese read
 * nine months of twelve — one column, half formatted and half raw, which is
 * the rendering this file exists to prevent. Widening the class widens what
 * `dateShape` CLAIMS, and it is kept to Latin for that reason: every script it
 * admits is a script whose names are not in the table.
 *
 * The transliterations are still deliberately NOT listed. `fevrier` is not how
 * French spells it, and a table of guesses at how somebody's exporter mangles
 * a name is the open rule this file removed. What replaced them is the real
 * spelling, not a second guess at it.
 */
function monthFromName(word: string): number | undefined {
  // `hasOwnProperty`, because `word` is a cell out of the user's data and the
  // table is an object. `Object.freeze` does not remove the prototype chain:
  // a cell reading `1 constructor 2026` passes the shape gate — `NAMED_DATE`
  // takes any word of three letters or more — and `MONTH_NAMES["constructor"]`
  // then answers the `Object` FUNCTION, which `known !== undefined` accepts as
  // a month. `__proto__` answers `Object.prototype` the same way.
  //
  // Today both come out benign, and only by luck: a function reaches
  // `Date.UTC` as a month, the arithmetic is NaN, the date is invalid, and the
  // cell is printed as it stands — which is the right answer arrived at by
  // accident. This repo's own rule is to guard any table keyed by a config or
  // data string, and the guard is one line where the luck is one refactor deep.
  //
  // NFC before the lookup, because `MONTH_LETTER` admits a combining mark and
  // the table is written with precomposed letters. A CSV exported on macOS
  // carries `f\u00e9vrier` as `e` + U+0301, which is a DIFFERENT string from the
  // key here and identical to it on screen and in a diff. Without this the
  // widened character class types the column a date and then refuses every
  // accented row in it, which is the exact rendering this file exists to
  // prevent, reintroduced by the change that was meant to end it.
  const key = word.normalize("NFC").toLowerCase();
  return Object.prototype.hasOwnProperty.call(MONTH_NAMES, key) ? MONTH_NAMES[key] : undefined;
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
  // The FIRST colon splits, and everything after it is the argument.
  // `split(":", 2)` truncates rather than keeping the remainder, so
  // `date:yyyy-MM-dd 00:00` handed `formatDate` the pattern `yyyy-MM-dd 00`
  // and printed a date with half a time after it. A date pattern is free text
  // with tokens in it and may hold any punctuation the author wants.
  const cut = spec.indexOf(":");
  const kind = cut < 0 ? spec : spec.slice(0, cut);
  const arg = cut < 0 ? "" : spec.slice(cut + 1);
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
