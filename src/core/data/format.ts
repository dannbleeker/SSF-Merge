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
 * A two-digit year, in the window the application it was copied from uses.
 *
 * Excel reads 00-29 as the 2000s and 30-99 as the 1900s, and a cell formatted
 * `dd/mm/yy` is ordinary there. This added 2000 to everything, so `15/06/85`
 * merged as 2085 — a birth date or a contract history a century out on every
 * slide, silently, with the column still typed a clean `date`.
 *
 * The window is Excel's rather than invented here, because the cell came out of
 * Excel and a merge that reads it differently from the sheet is the surprise.
 */
function fullYear(year: number): number {
  if (year >= 100) return year;
  return year < 30 ? 2000 + year : 1900 + year;
}

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
// The tail is a TIME, not anything. `.*` accepted any text after the day, so a
// PERIOD — `2026-03-01 - 2026-03-31`, an ordinary way to write a range — was
// typed as a date and formatted as its first half, with the rest discarded and
// nothing to say so. A cell this engine cannot fully claim is printed as the
// author typed it; that is the rule every other refusal here follows.
const ISO_DATE = /^\d{4}-\d{2}-\d{2}(?:[T ]\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/;
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

/**
 * A number written with thousands groups: `1,234`, `1.234.567`, `-12,345`.
 *
 * The FIRST group may not be zero, and that is the whole of the difference
 * between an ambiguous cell and a wrong answer. `1,500` is genuinely ambiguous
 * — fifteen hundred or one-and-a-half — and this engine reads it as grouping,
 * which is decided and written down. `0,500` is not ambiguous at all: no
 * spreadsheet writes a leading `0` group for five hundred, so it can only be
 * the decimal five tenths. It was read as 500.
 *
 * That is a factor of a thousand on ordinary Danish and German data — a rate,
 * a share, a percentage shown to three decimals — and the same column could
 * hold both readings at once: `0,250` came back 250 while `0,05` came back
 * 0.05, one column, typed `number`, two answers a thousandfold apart. The value
 * also goes straight into a chart, so the bar plots 250 where the sheet says a
 * quarter.
 */
const GROUPED_COMMA = /^-?[1-9]\d{0,2}(?:,\d{3})+$/;

/** The same rule for the European spelling: `1.234`, `12.345.678`. */
const GROUPED_DOT = /^-?[1-9]\d{0,2}(?:\.\d{3})+$/;

/** Parse the number forms a spreadsheet actually produces, including the European one. */
export function numericValue(raw: string): number | undefined {
  const text = numericText(raw);
  if (text === undefined) return undefined;
  const n = Number(text);
  return Number.isFinite(n) ? n : undefined;
}

/**
 * The plain decimal a cell spells, with its grouping resolved — the string
 * `numericValue` parses.
 *
 * Exposed because the DIGITS matter to one caller and the double does not. A
 * cell can hold more significant digits than a double carries, and the only way
 * to know whether formatting changed the value is to compare what was written
 * with what came back. See `applyFormat`.
 */
function numericText(raw: string): string | undefined {
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
    normalised = GROUPED_COMMA.test(normalised) ? normalised.replace(/,/g, "") : normalised.replace(",", ".");
  } else if (hasDot && GROUPED_DOT.test(normalised)) {
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
  return Number.isFinite(Number(normalised)) ? normalised : undefined;
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
    return utcDate(fullYear(year), month, day);
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
    return utcDate(fullYear(year), month, Number(named[1]));
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
  // The SHORT forms for the six languages that were named in the manual and
  // whose abbreviations this table did not hold. Ten languages were claimed;
  // only the English and Scandinavian abbreviations were here, so a German
  // column of `1 dez 2026` typed as a date on shape and then printed as the raw
  // cell — a merged deck with some slides formatted and some not, and nothing
  // saying why. That is precisely the pathology `dateShape` exists to avoid,
  // arriving through the table rather than around it.
  //
  // Every word below maps to the month it means in EVERY language that writes
  // it — `mar` is March in English, Spanish and Portuguese; `dez` is December
  // in German and Portuguese; `set` is September in Italian and Portuguese —
  // which is the property that lets one table serve them all, and it is
  // asserted over the whole table rather than trusted.
  //
  // German. `Mrz` as well as `M\u00e4r`: CLDR says `M\u00e4r` and Excel on
  // Windows writes `Mrz`, and it is the spreadsheet's spelling that arrives in
  // a pasted column. A table built from the standard alone reads eleven of a
  // German year's twelve months, which is the half-formatted deck this table
  // exists to prevent, missing by one word.
  "m\u00e4r": 3,
  mrz: 3,
  dez: 12,
  // Dutch
  mrt: 3,
  // French. Four letters where the language writes four: `janv`, `f\u00e9vr`,
  // `juil` and `sept` are what a French spreadsheet exports, and a three-letter
  // truncation of them is not a form anybody writes.
  janv: 1,
  "f\u00e9vr": 2,
  avr: 4,
  juil: 7,
  sept: 9,
  "d\u00e9c": 12,
  // Spanish. `may` is the English full name and is already above — the two
  // spell May the same way, which is the shared-word property this table rests
  // on rather than an omission.
  ene: 1,
  abr: 4,
  ago: 8,
  dic: 12,
  // Italian
  gen: 1,
  mag: 5,
  giu: 6,
  lug: 7,
  set: 9,
  ott: 10,
  // Portuguese. `out` is an English word as well as October in Portuguese, and
  // it is safe here because `NAMED_DATE` admits a month word only between a
  // one-or-two-digit number and a two-to-four-digit year: "2 out 3" is not a
  // date shape and never reaches this table.
  fev: 2,
  out: 10,
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
 * ONE PASS over the pattern's letter runs, rather than seven `replace` calls.
 * The chain it replaces went wrong twice in opposite directions.
 *
 * Unbounded, a token matched inside an ordinary word: `Odds: d MMM yyyy`
 * printed `O01s: 1 Mar 2026`, and `dddd` — the Excel and .NET spelling of a
 * weekday name, which a template author reaches for — printed `0101`.
 *
 * Bounded against any letter, a pattern made only of tokens stopped matching
 * at all: `yyyyMMdd`, the ISO-basic spelling used for file names, printed
 * itself, and every token in it is one the manual lists. A pattern of nothing
 * but tokens printed as written is precisely what the manual says will not
 * happen.
 *
 * So the unit is the LETTER RUN. A run is read greedily, longest token first —
 * `yyyy` before `yy`, `MMMM` before `MMM` before `MM`, `dd` before `d`, because
 * reordering them turns `yyyy` into `2626` — and:
 *
 * - a run is tokenized only if the WHOLE of it is tokens, which is what keeps
 *   `Odds`, `address`, `Wedding`, `den`, `dato` and `due` intact;
 * - two adjacent tokens spelling the same letter make the run literal, which is
 *   what keeps `dddd` a weekday name rather than two days of the month. (Only
 *   the EXACT multiples need it: `dddd` is `dd` twice and `yyyyyy` is `yy`
 *   three times, where `MMMMM` and `yyy` are already refused by the whole-run
 *   rule because their last letter matches no token. This paragraph used to
 *   credit `MMMMM` to the same-letter check; removing that check and running it
 *   says otherwise.)
 *
 * "Starts with a token, tail printed as written" was tried and is worse. It
 * reads `yyyy-MM-ddTHH:mm` as a date with `THH:mm` after it — which is nice —
 * and it also turns `Berlin, den d. MMMM yyyy` into `Berlin, 1en 1. March
 * 2026`, because `den` starts with a `d`. `den` is the ordinary Danish, German
 * and Swedish long-date word and the manual invites literal text in the
 * pattern, so a rule that eats it is not a rule this can have. `H` is not a
 * token here and never was; a run holding one is text.
 *
 * No lookbehind, deliberately. The seven-`replace` version was this file's only
 * use of one, and a lookbehind in a regex LITERAL is a parse-time SyntaxError
 * on WebKit below Safari 16.4 — which is a module that fails to load, i.e. a
 * blank pane, rather than a date printed wrong.
 */
const DATE_TOKENS: readonly [token: string, letter: string, value: (d: Date) => string][] = [
  ["yyyy", "y", (d) => String(d.getUTCFullYear())],
  ["yy", "y", (d) => pad(d.getUTCFullYear() % 100)],
  ["MMMM", "M", (d) => MONTHS_FULL_EN[d.getUTCMonth()] ?? ""],
  ["MMM", "M", (d) => MONTHS_EN[d.getUTCMonth()] ?? ""],
  ["MM", "M", (d) => pad(d.getUTCMonth() + 1)],
  ["dd", "d", (d) => pad(d.getUTCDate())],
  ["d", "d", (d) => String(d.getUTCDate())],
];

/** One run of ASCII letters, tokenized — or nothing, meaning print it as written. */
function tokenizeRun(run: string, d: Date): string | undefined {
  let out = "";
  let at = 0;
  let previous: string | undefined;
  for (;;) {
    const hit = DATE_TOKENS.find(([token]) => run.startsWith(token, at));
    if (!hit) break;
    // `dddd` is `dd` twice, and it is a weekday name rather than the day of
    // the month written out. Same for `yyyyyy`. A run that repeats a token's
    // letter is not a pattern this add-in knows.
    if (hit[1] === previous) return undefined;
    out += hit[2](d);
    previous = hit[1];
    at += hit[0].length;
  }
  // The whole run, or none of it. A run that is part token and part word is a
  // word: `den`, `dato`, `due` and `deadline` all begin with a token's letter,
  // and the manual promises that text in a pattern prints as written.
  if (at !== run.length) return undefined;
  return out;
}

/**
 * A WORD, for the purpose of the rule above: a run of LATIN letters, with the
 * combining marks and apostrophes that sit inside one.
 *
 * ASCII alone cut a word at its own accent, and a one-character run of `d` is a
 * whole token — so `día d de MMMM` printed `1ía 1 de March` and
 * `Date d'échéance dd/MM/yyyy` printed `Date 1'échéance 01/03/2026`.
 *
 * ANY Unicode letter was the first fix and it is worse, because the tokens are
 * Latin and every other script's date words are not: `yyyy年MM月dd日` — the
 * ordinary Japanese and Chinese pattern — became one run holding `年`, which is
 * no token, so the whole thing printed itself. So did `yyyy년 MM월 dd일` and
 * `dd MMMM yyyyг.`. The manual invites writing the month in your own language,
 * and that broke it for every language not written in this alphabet.
 *
 * Latin is the line because it is where the tokens live. A letter from another
 * script cannot be part of a word a token is hiding inside, so it does not need
 * to end one.
 *
 * `\p{Mn}` because a decomposed `é` is `e` plus a combining mark, which is not
 * a letter — text pasted from a Mac routinely arrives that way. Both
 * apostrophes because U+2019 is what Word's autocorrect produces and U+0027 is
 * what a keyboard does.
 */
const PATTERN_WORD = /[\p{Script=Latin}\p{Mn}]+(?:['\u2019][\p{Script=Latin}\p{Mn}]+)*/gu;

const APOSTROPHE = /(['\u2019])/;

/**
 * A word run, formatted — the whole run if it is all tokens, else each
 * apostrophe-separated piece if EVERY piece is.
 *
 * The second rule is not decoration. Holding an apostrophe inside a word is
 * what stops `Date d'échéance` printing as `Date 1'échéance`, and it welds the
 * two halves of `MMM'yy` into one run that is not a token — so the commonest
 * abbreviated pattern there is printed itself. `MMM'yy` is `Mar'26` and
 * `MMM’yy` is `Mar’26` — the apostrophe is kept as written. `d'échéance` is
 * left alone, because `échéance` is not a token and
 * one piece failing fails the run.
 *
 * The price is the reverse case: `d'MMMM` is `1'March` rather than itself. Both
 * readings are guesses about a pattern nobody writes, and the rule that gets
 * `MMM'yy` right is the one worth having.
 */
function formatWord(run: string, d: Date): string {
  const whole = tokenizeRun(run, d);
  if (whole !== undefined) return whole;
  const pieces = run.split(APOSTROPHE);
  if (pieces.length === 1) return run;
  const out: string[] = [];
  for (let i = 0; i < pieces.length; i += 2) {
    const piece = pieces[i] ?? "";
    const formatted = piece === "" ? undefined : tokenizeRun(piece, d);
    if (formatted === undefined) return run;
    out.push(formatted);
    const separator = pieces[i + 1];
    if (separator !== undefined) out.push(separator);
  }
  return out.join("");
}

export function formatDate(d: Date, pattern: string): string {
  return pattern.replace(PATTERN_WORD, (run) => formatWord(run, d));
}

/**
 * `Math.abs(n)` at `decimals` places, rounded the way a spreadsheet rounds.
 *
 * NOT `toFixed`, which rounds the binary double rather than the decimal the
 * author typed. `1.005` is stored a hair below 1.005, so `toFixed(2)` answers
 * "1.00" where Excel — and the cell the number was copied out of — says 1.01.
 * Comparing the two over a corpus of ordinary money and percentage values put
 * them apart on 118 of 310, which is not a corner: it is every value whose last
 * kept digit is followed by a 5.
 *
 * So the rounding is done on the DECIMAL string, half away from zero. The
 * string is JavaScript's shortest round-trip spelling, which is the one that
 * reads back as the number the author typed — for `1.005` it is "1.005", and
 * the decision is then arithmetic anybody can check by eye rather than a
 * property of the binary value underneath.
 *
 * A magnitude with no plain spelling — above 1e21, where JavaScript's shortest
 * form is exponential — has no fixed-point digits to round or to group, and is
 * handled in `formatNumber` rather than here.
 */
function fixedDecimal(n: number, decimals: number): string {
  const plain = String(Math.abs(n));
  if (plain.includes("e") || plain.includes("E")) return Math.abs(n).toFixed(decimals);
  return roundDigits(plain, decimals);
}

/**
 * The same rounding, on a decimal string that came from somewhere else.
 *
 * `fixedDecimal` rounds the DOUBLE's shortest spelling, which is the right
 * input when a number is all anybody has. It is the wrong one when the cell is
 * still in hand and carries more digits than the double: `2.4999999999999999999`
 * is stored as 2.5, so the double's spelling rounds UP to 3 where the cell
 * rounds down to 2 — a different number on the slide, grouped and formatted as
 * though it were right.
 *
 * `applyFormat` therefore rounds the cell's own digits. The magnitude has
 * already been checked against the double (`sameNumber`), so this is about the
 * last printed digit and nothing else.
 *
 * `plain` is unsigned, with no exponent: `numericText`'s output, less its sign.
 */
function roundDigits(plain: string, decimals: number): string {
  const [whole = "0", frac = ""] = plain.split(".");
  if (frac.length <= decimals) return decimals === 0 ? whole : `${whole}.${frac.padEnd(decimals, "0")}`;
  const kept = (whole + frac.slice(0, decimals)).split("");
  // Half AWAY from zero, which is what a spreadsheet does and what somebody
  // reading the cell expects. Banker's rounding would answer 1.00 here and be
  // defensible; it would also disagree with the number on screen in Excel,
  // which is the comparison a reader of the slide actually makes.
  if ((frac.charCodeAt(decimals) ?? 0) >= 53) {
    let i = kept.length - 1;
    for (; i >= 0; i--) {
      if (kept[i] === "9") kept[i] = "0";
      else {
        kept[i] = String(Number(kept[i]) + 1);
        break;
      }
    }
    // Every digit was a nine: 9.99 at one place is 10.0, and the carry needs a
    // column that was not there.
    if (i < 0) kept.unshift("1");
  }
  const digits = kept.join("");
  const cut = digits.length - decimals;
  return decimals === 0 ? digits : `${digits.slice(0, cut) || "0"}.${digits.slice(cut)}`;
}

/**
 * Whether the digits this format will PRINT survived the parse.
 *
 * The comparison is on the INTEGER part alone, and that is the whole subtlety.
 * Grouping prints every integer digit, so one lost there is a different number
 * on the slide — `1234567890123456789` came back as `1 234 567 890 123 456
 * 800`, an order number nobody typed. The fraction is rounded to `decimals`
 * places before it is printed, so a cell carrying more precision than a double
 * holds is not a defect there: `0.3333333333333333333` at two places is `0,33`
 * either way.
 *
 * Comparing the WHOLE decimal was the first rule and it refused those — a
 * Danish deck then printed `1234567890.12345678` with a dot, because the raw
 * fall-through keeps the cell's own separator. It also refused every value
 * JavaScript spells with an exponent, including small ones like `0.0000001`
 * that group perfectly.
 *
 * Text, not arithmetic: `Number(wrote) === n` is always true, because the
 * digits that were lost were lost in the parse.
 */
function sameNumber(wrote: string, n: number): boolean {
  const spelled = String(Math.trunc(n));
  // A magnitude with no plain integer spelling has nothing to compare and
  // nothing to group; `applyFormat` refuses it a line later for that reason.
  if (spelled.includes("e") || spelled.includes("E")) return false;
  const whole = (s: string): string => {
    const body = (s.replace(/^[+-]/, "").split(".")[0] ?? "").replace(/^0+(?=\d)/, "");
    return body === "" ? "0" : body;
  };
  return whole(wrote) === whole(spelled);
}

/**
 * Group a rounded decimal string and give it its sign.
 *
 * Shared by the public `formatNumber` and by `applyFormat`, which rounds the
 * CELL rather than the double and so cannot go through the first. The two used
 * to be one function, and separating the rounding from the grouping is what let
 * the cell's digits reach the slide.
 */
function groupFixed(fixed: string, negative: boolean, group: string, point: string): string {
  const [whole = "0", frac] = fixed.split(".");
  const grouped = whole.replace(/\B(?=(\d{3})+(?!\d))/g, group);
  // The sign comes from the ROUNDED value, not the input. `-0.4` at no decimal
  // places rounds to zero, and taking the sign from the input printed it as
  // `-0` — a quantity that does not exist, on a slide, from an ordinary cell.
  return `${negative && Number(fixed) !== 0 ? "-" : ""}${grouped}${frac ? point + frac : ""}`;
}

export function formatNumber(n: number, decimals: number, group: string, point: string): string {
  // Nothing to group. Above 1e21 JavaScript spells a number with an exponent,
  // and `1.2345678901234568e+24` split on its dot and grouped comes out as
  // `1,2345678901234568e+24` — a European decimal, on a slide, from a whole
  // number. `applyFormat` returns the cell unchanged before it gets here, which
  // is its own contract; this is the same defect reached through the PUBLIC
  // export, which has no cell to fall back to and answers the value as
  // JavaScript spells it.
  //
  // The digits-changed test that `applyFormat` applies CANNOT be made here: it
  // compares the value against the cell it was read from, and this entry point
  // has no cell.
  if (!Number.isFinite(n) || Math.abs(n) >= 1e21) return String(n);
  return groupFixed(fixedDecimal(n, decimals), n < 0, group, point);
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
      // A cell the double CHANGED is returned unchanged, which is this
      // function's own contract for a value that does not match its format.
      //
      // Two shapes. Above 1e21 JavaScript spells a number with an exponent, so
      // `1e21` printed "1e+21" and a 25-digit cell printed
      // "1,2345678901234568e+24" — a European decimal, on a slide, from an
      // integer. Below that the spelling is fine and the DIGITS can be wrong:
      // `1234567890123456789` came out as `1 234 567 890 123 456 800`, which is
      // not the number in the cell, printed with the confidence of a formatted
      // one. An order number, an IBAN typed without spaces and a 19-digit
      // identifier all land there.
      //
      // The test is the CELL against the value, not a bound on the value.
      // `Number.isSafeInteger` was the first rule and it refuses exact
      // magnitudes too — 2^53 itself, `1e18`, `1e20` — which a double carries
      // perfectly and which a user has every right to see grouped. Compare what
      // was written with what came back and refuse only a disagreement.
      // `numericText` answers for exactly the cells `numericValue` does, and
      // that has already answered, so this is the same string it parsed.
      const wrote = numericText(raw) ?? "";
      if (!sameNumber(wrote, n)) return raw;
      // Rounded from the CELL, not from the double. `2.4999999999999999999` is
      // stored as 2.5, so rounding the double's spelling answers 3 where the
      // cell answers 2 — a different number on the slide, grouped and formatted
      // as though it were right. The magnitude has already been checked against
      // the double; this is the last printed digit.
      // CANONICALISED, not merely unsigned. `numericText` preserves whatever the
      // cell spelled, and `0007` is a number shape — so the zeros reached
      // `groupFixed` and were thousands-grouped: `0000000001234` printed as
      // `0 000 000 001 234`. A zero-padded column of order numbers, article
      // codes or postcodes types as a number, so this is an ordinary path.
      // Going through `formatNumber` used to canonicalise for free, because it
      // starts from `String(Math.abs(n))`; rounding the cell means doing it
      // here. `sameNumber` already strips them before comparing, which is why
      // the guard above let these through.
      const digits = wrote.replace(/^[+-]/, "").replace(/^0+(?=\d)/, "");
      return groupFixed(roundDigits(digits, decimals), n < 0, " ", ",");
    }
    case "date": {
      const d = parseDate(raw);
      return d ? formatDate(d, arg || "dd-MM-yyyy") : raw;
    }
    default:
      return raw;
  }
}
