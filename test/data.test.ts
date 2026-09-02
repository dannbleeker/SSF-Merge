import { describe, expect, it } from "vitest";
import { applyFormat, formatDate, formatNumber, numericValue, parseDate } from "../src/core/data/format.js";
import { detectType, looksLikeDate, looksLikeNumber, parseDelimited, toRecordSet } from "../src/core/data/recordset.js";

describe("parseDelimited", () => {
  it("reads a range pasted out of Excel, which is tab separated", () => {
    expect(parseDelimited("Name\tRegion\nAda\tNordics")).toEqual([
      ["Name", "Region"],
      ["Ada", "Nordics"],
    ]);
  });

  it("keeps a comma that is inside a quoted company name", () => {
    expect(parseDelimited('Name,Note\n"Contoso, Ltd",fine')).toEqual([
      ["Name", "Note"],
      ["Contoso, Ltd", "fine"],
    ]);
  });

  it("reads a doubled quote as one quote", () => {
    expect(parseDelimited('a\n"say ""hi"""')).toEqual([["a"], ['say "hi"']]);
  });

  it("leaves a ONE-COLUMN paste in one column, whatever it contains", () => {
    /**
     * `chooseDelimiter` scores each candidate and correctly finds that none of
     * them separates anything here — and then fell back to the comma, which
     * `splitOn` duly split on. A column of European decimals came apart:
     * `1,5` became `1` and `5`, and every slide printed `1`. A column of
     * company names did the same to `Contoso, Ltd`.
     *
     * Its own docstring already said what should happen — "a one-column paste
     * has no separator to find, and every caller wants one column rather than a
     * refusal" — and the fallback defeated it by choosing a character that is
     * in the data. The asymmetry was visible in this file: the semicolon case
     * was pinned as working while the identical comma case was not.
     */
    expect(parseDelimited("Bel\u00f8b\n1,5\n2,5")).toEqual([["Bel\u00f8b"], ["1,5"], ["2,5"]]);
    expect(parseDelimited("Kunde\nContoso, Ltd")).toEqual([["Kunde"], ["Contoso, Ltd"]]);
    expect(parseDelimited("Notes; extra\nfine\nalso")).toEqual([["Notes; extra"], ["fine"], ["also"]]);
    // And a real comma-separated table is still read as one.
    expect(parseDelimited("Name,Region\nAda,Nordics")).toEqual([
      ["Name", "Region"],
      ["Ada", "Nordics"],
    ]);
  });
});

describe("toRecordSet", () => {
  it("names an unnamed column instead of dropping it", () => {
    const rs = toRecordSet([
      ["Name", ""],
      ["Ada", "x"],
    ]);
    expect(rs.columns.map((c) => c.name)).toEqual(["Name", "Column 2"]);
    expect(rs.rows[0]?.["Column 2"]).toBe("x");
  });

  it("keeps both of two columns with the same header", () => {
    const rs = toRecordSet([
      ["Name", "Name"],
      ["a", "b"],
    ]);
    expect(rs.columns.map((c) => c.name)).toEqual(["Name", "Name 2"]);
  });

  it("drops a wholly blank row and keeps a partly filled one", () => {
    const rs = toRecordSet([
      ["A", "B"],
      ["", ""],
      ["x", ""],
    ]);
    expect(rs.rows).toHaveLength(1);
  });

  it("cannot be reached through the prototype chain", () => {
    // A column called __proto__ is a legal spreadsheet header.
    const rs = toRecordSet([["__proto__"], ["x"]]);
    expect(rs.rows[0]?.["__proto__"]).toBe("x");
  });
});

describe("detectType", () => {
  it("calls a column of numbers a number", () => {
    expect(detectType(["1", "2 000", "3,5"])).toBe("number");
  });

  it("calls a column with one word in it text", () => {
    expect(detectType(["1", "n/a"])).toBe("text");
  });

  it("calls an ISO date column a date", () => {
    expect(detectType(["2026-03-01", "2026-04-02"])).toBe("date");
  });
});

describe("looksLikeDate", () => {
  it("refuses a slash date where both numbers could be a month", () => {
    // 03/01/2026 is 3 January in Copenhagen and 1 March in Chicago. A deck that
    // draws perfectly and is two months wrong is the worse outcome.
    expect(looksLikeDate("03/01/2026")).toBe(false);
  });

  it("accepts a slash date only one reading fits", () => {
    expect(looksLikeDate("15/01/2026")).toBe(true);
  });

  it("accepts ISO and a named month", () => {
    expect(looksLikeDate("2026-03-01")).toBe(true);
    expect(looksLikeDate("3 March 2026")).toBe(true);
  });
});

describe("numericValue", () => {
  it("reads a European decimal", () => {
    expect(numericValue("1,5")).toBe(1.5);
  });

  it("reads an American thousands group", () => {
    expect(numericValue("1,500")).toBe(1500);
  });

  it("reads a LEADING ZERO group as a decimal, because no other reading exists", () => {
    /**
     * `1,500` is genuinely ambiguous — fifteen hundred or one-and-a-half — and
     * this engine reads it as grouping, which is decided and written down.
     * `0,500` is not ambiguous at all: no spreadsheet writes a leading `0`
     * group for five hundred. It was read as 500.
     *
     * A factor of a thousand on ordinary Danish and German data — a rate, a
     * share, a percentage shown to three decimals — and one column could hold
     * both readings at once, which is the pathology this file's own docstrings
     * say they exist to prevent. The value goes straight into charts too, so
     * the bar plotted 250 where the sheet said a quarter.
     */
    expect(numericValue("0,500")).toBe(0.5);
    expect(numericValue("0.500")).toBe(0.5);
    expect(numericValue("-0,004")).toBeCloseTo(-0.004);
    // The whole column reads one way now.
    expect(numericValue("0,250")).toBe(0.25);
    expect(numericValue("0,05")).toBe(0.05);
    // And the ambiguous case is untouched, which is the decision that stands.
    expect(numericValue("1,500")).toBe(1500);
    expect(numericValue("1.234")).toBe(1234);
  });

  it("reads a mixed form by whichever separator comes last", () => {
    expect(numericValue("1.234,56")).toBeCloseTo(1234.56);
    expect(numericValue("1,234.56")).toBeCloseTo(1234.56);
  });

  it("answers undefined for something that is not a number", () => {
    expect(numericValue("n/a")).toBeUndefined();
  });
});

describe("applyFormat", () => {
  it("groups a number and keeps the sign", () => {
    expect(formatNumber(-1234567.5, 2, " ", ",")).toBe("-1 234 567,50");
  });

  it("rounds the number the author typed, not the double underneath it", () => {
    /**
     * `toFixed` rounds the binary value. `1.005` is stored a hair below 1.005,
     * so it answers "1.00" where Excel — and the cell the number was copied out
     * of — says 1.01. Over a corpus of ordinary money and percentage values the
     * two disagreed on 118 of 310, which is not a corner: it is every value
     * whose last kept digit is followed by a 5.
     *
     * A slide that disagrees with the spreadsheet beside it about a price is
     * the one arithmetic error nobody blames on the tool.
     */
    expect(formatNumber(1.005, 2, " ", ","), "toFixed says 1.00").toBe("1,01");
    expect(formatNumber(2.675, 2, " ", ","), "toFixed says 2.67").toBe("2,68");
    expect(formatNumber(-1.005, 2, " ", ",")).toBe("-1,01");
    // Half AWAY from zero, both directions, which is what a spreadsheet does.
    expect(formatNumber(2.5, 0, " ", ",")).toBe("3");
    expect(formatNumber(-2.5, 0, " ", ",")).toBe("-3");
  });

  it("carries a rounding that runs out of columns", () => {
    // 9.99 at one place is 10.0, and the carry needs a digit that was not
    // there. All-nines is the case a carry loop gets wrong.
    expect(formatNumber(9.99, 1, " ", ",")).toBe("10,0");
    expect(formatNumber(0.999, 2, " ", ",")).toBe("1,00");
    expect(formatNumber(99.995, 2, " ", ",")).toBe("100,00");
    expect(formatNumber(999.5, 0, " ", ",")).toBe("1 000");
  });

  it("leaves a number too large to spell in full exactly as the cell has it", () => {
    /**
     * Above 1e21 JavaScript spells a number with an exponent, and there is no
     * fixed-point form to group. `1e21` printed "1e+21"; a 25-digit integer
     * printed "1,2345678901234568e+24" — a European decimal, on a slide, from
     * a whole number. Unchanged is this function's own contract for a value
     * that does not match its format.
     */
    expect(applyFormat("1000000000000000000000", "number:0")).toBe("1000000000000000000000");
    // And BELOW 1e21, where the spelling is fine and the digits are not. A
    // double cannot carry nineteen significant digits, so this printed
    // "1 234 567 890 123 456 800" — not the number in the cell, with the
    // confidence of a formatted one. An order number or a 19-digit identifier
    // lands here, and changing one silently is worse than leaving it alone.
    expect(applyFormat("1234567890123456789", "number:0")).toBe("1234567890123456789");
    expect(applyFormat("9007199254740993", "number:0")).toBe("9007199254740993");
    // And an exactly-representable magnitude still formats, however large. The
    // first rule was `Number.isSafeInteger`, which refuses 2^53 itself and
    // every exact power above it — a bound on the VALUE where the question is
    // whether the double changed the CELL. Compare what was written with what
    // came back, and refuse only a disagreement.
    expect(applyFormat("9007199254740991", "number:0")).toBe("9 007 199 254 740 991");
    expect(applyFormat("9007199254740992", "number:0")).toBe("9 007 199 254 740 992");
    expect(applyFormat("-9007199254740992", "number:0")).toBe("-9 007 199 254 740 992");
    expect(applyFormat("1000000000000000000", "number:0")).toBe("1 000 000 000 000 000 000");
    expect(applyFormat("100000000000000000000", "number:0")).toBe("100 000 000 000 000 000 000");
    // A grouped cell still reads, and a fraction still rounds.
    expect(applyFormat("1.234.567", "number:0")).toBe("1 234 567");
    expect(applyFormat("12345678901234.5", "number:0")).toBe("12 345 678 901 235");
    // The comparison is on the INTEGER digits, which are the ones grouping
    // prints in full. A cell carrying more PRECISION than a double holds is not
    // a defect: the fraction is rounded to the places asked for either way.
    // Comparing the whole decimal refused these — and the raw fall-through
    // keeps the cell's own separator, so a Danish deck printed a dot.
    expect(applyFormat("1234567890.12345678", "number:2")).toBe("1 234 567 890,12");
    expect(applyFormat("0.3333333333333333333", "number:2")).toBe("0,33");
    expect(applyFormat("2,5000000000000001", "number:2")).toBe("2,50");
    // Including a magnitude JavaScript spells with an exponent at the SMALL
    // end, which groups perfectly.
    expect(applyFormat("0.0000001", "number:2")).toBe("0,00");
    // And the ROUNDING is the cell's, not the double's. `2.4999999999999999999`
    // is stored as 2.5, so rounding the double's spelling answers 3 where the
    // cell answers 2 — a different number on the slide, grouped and formatted
    // as though it were right. Measured at 814 wrong in 20 000 thirteen-digit
    // cells before this; zero after.
    expect(applyFormat("2.4999999999999999999", "number:0")).toBe("2");
    expect(applyFormat("0.1249999999999999999", "number:2")).toBe("0,12");
    expect(applyFormat("1.0049999999999999999", "number:2")).toBe("1,00");
    // The ordinary half-away-from-zero cases still round up, which is the
    // behaviour a spreadsheet shows and the reason this file rounds strings.
    // (`1.0050` rather than `1.005`, which is a THOUSANDS group by this file's
    // own rule — a lone three-digit run after a separator.)
    expect(applyFormat("1.0050", "number:2")).toBe("1,01");
    expect(applyFormat("-0.4", "number:0")).toBe("0");
    expect(applyFormat("-1.5", "number:0")).toBe("-2");
    expect(applyFormat("1234567890123456789012345", "number:2")).toBe("1234567890123456789012345");
    // And the ordinary case is untouched.
    expect(applyFormat("1234567.891", "number:2")).toBe("1 234 567,89");
    // The PUBLIC formatter has no cell to fall back to, and must not mangle it
    // either: it answered "1,2345678901234568e+24", a European decimal made out
    // of an exponent.
    expect(formatNumber(1.234_567_890_123_456_8e24, 2, " ", ",")).toBe("1.2345678901234568e+24");
    expect(formatNumber(Number.POSITIVE_INFINITY, 2, " ", ",")).toBe("Infinity");
  });

  it("formats a date to the pattern in the template", () => {
    expect(applyFormat("2026-03-01", "date:dd MMM yyyy")).toBe("01 Mar 2026");
  });

  it("keeps a date pattern that has a colon in it", () => {
    // `split(":", 2)` TRUNCATES rather than putting the remainder in the last
    // element, so everything after the second colon was dropped: a template
    // asking for `date:yyyy-MM-dd 00:00` — a date with a fixed time after it,
    // which is an ordinary thing to write on a slide — printed
    // "2026-03-01 00" and stopped. A date pattern is free text with tokens in
    // it and may hold any punctuation its author wants.
    expect(applyFormat("2026-03-01", "date:yyyy-MM-dd 00:00")).toBe("2026-03-01 00:00");
    expect(applyFormat("2026-03-01", "date:dd/MM/yyyy 12:30:00")).toBe("01/03/2026 12:30:00");
  });

  it("returns the cell unchanged when it does not match its format", () => {
    // The cell is what the user typed. Showing it beats showing nothing.
    expect(applyFormat("not a number", "number:2")).toBe("not a number");
  });

  it("leaves an ambiguous date alone rather than picking a reading", () => {
    expect(applyFormat("03/01/2026", "date:yyyy-MM-dd")).toBe("03/01/2026");
  });

  it("upper-cases with the locale, not the ASCII table", () => {
    expect(applyFormat("måned", "upper")).toBe("MÅNED");
  });

  it("parses an unambiguous slash date day-first", () => {
    expect(parseDate("15/01/2026")?.toISOString().slice(0, 10)).toBe("2026-01-15");
  });

  it("refuses every ambiguous pair, including the two at the edge of the year", () => {
    // The refusal is `a <= 12 && b <= 12 && a !== b`, and each of those three
    // clauses owns a case nothing was checking. `scripts/mutate-core.mjs` found
    // it: `<= 12` loosened to `< 12` left the suite green, so every DECEMBER
    // ambiguity — the commonest month in a renewal column — would have started
    // parsing on one reading with nothing to say which.
    //
    // The owner decided this on 2026-08-12 with the cost on the table: a US
    // author writing 03/01/2026 for 1 March gets an empty cell, because the
    // alternative is a Gantt that draws perfectly and is two months wrong.
    for (const d of ["03/01/2026", "12/01/2026", "01/12/2026"]) {
      expect(parseDate(d), `${d} could be either reading`).toBeUndefined();
    }
    // And the two that are NOT ambiguous, which is what stops the rule
    // swallowing dates it has no business refusing: same number both sides, and
    // a first number no month could be.
    expect(parseDate("12/12/2026")?.toISOString().slice(0, 10)).toBe("2026-12-12");
    expect(parseDate("13/01/2026")?.toISOString().slice(0, 10)).toBe("2026-01-13");
  });

  it("reads a number Excel grouped with the space it actually uses", () => {
    /**
     * A number formatted with a SPACE group — Swedish, Norwegian, Finnish,
     * French, Polish, Czech, Russian — is not written with U+0020. Excel uses
     * a NO-BREAK SPACE, and modern builds use the NARROW no-break space in
     * French, precisely so the number cannot be broken across a line. Copying
     * that cell puts the DISPLAYED text on the clipboard, so a paste into the
     * pane carries the space Excel chose and not the one a keyboard writes.
     *
     * `NUMBER` admitted `[ .,]` — plain ASCII — so the whole column typed as
     * text: `{{Revenue|number:2}}` left the cell exactly as pasted, and the
     * chart writer refused every value and counted it unfilled. Nothing said
     * why, because the cell LOOKS like a number and the pane's own output uses
     * a space too. The population that meets this is the same one the
     * semicolon delimiter was added for.
     */
    for (const space of ["\u00a0", "\u202f", "\u2009"]) {
      const grouped = `1${space}234${space}567`;
      expect(numericValue(grouped), JSON.stringify(space)).toBe(1234567);
      expect(looksLikeNumber(grouped), JSON.stringify(space)).toBe(true);
      expect(applyFormat(grouped, "number:0"), JSON.stringify(space)).toBe("1 234 567");
      // A decimal comma behind that grouping, which is the form the same
      // locales write and the reason the group and the decimal may not be the
      // same character.
      expect(numericValue(`1${space}234,5`), JSON.stringify(space)).toBe(1234.5);
    }
  });

  it("still refuses a number whose groups disagree, whichever spaces they are", () => {
    // The property `NUMBER`'s backreference holds, extended to the new
    // characters rather than left behind by them: the separator is captured
    // once and every later group must repeat THAT one. Mixing them is not a
    // locale, it is a paste that went through something.
    expect(numericValue("1\u00a0234\u202f567")).toBeUndefined();
    expect(numericValue("1.234\u00a0567")).toBeUndefined();
    // But a space group with a DOT decimal is a real number and always has
    // been with the ASCII space, so it stays one with the others. The
    // backreference forbids reusing the GROUP character, not every separator.
    expect(numericValue("1\u00a0234.567")).toBe(1234.567);
    expect(numericValue("1 234.567")).toBe(1234.567);
    // And a space is never the DECIMAL separator, so a trailing group of
    // anything but three digits is not a number.
    expect(numericValue("1\u00a05")).toBeUndefined();
  });

  it("reads a DATE the spreadsheet padded, in every spelling it accepts", () => {
    // A cell arrives with whatever whitespace its export gave it, and one
    // `.trim()` in `parseDate` is the whole of that. Removing it left the suite
    // green — so a padded renewal column would have been printed as written on
    // every slide, with no format applied and nothing to say why.
    //
    // The first version of this test asserted `numericValue` instead, and
    // passed against the mutation because that function trims somewhere the
    // sweep did not flag — which is a test proving a rule that was already
    // held. Checking WHICH assertion failed is what caught it.
    expect(parseDate(" 15/01/2026 ")?.toISOString().slice(0, 10)).toBe("2026-01-15");
    expect(parseDate("\t2026-03-01\n")?.toISOString().slice(0, 10)).toBe("2026-03-01");
    expect(parseDate(" 1 Mar 2026 ")?.toISOString().slice(0, 10)).toBe("2026-03-01");
    expect(applyFormat(" 15/01/2026 ", "date:yyyy-MM-dd")).toBe("2026-01-15");
  });
});

/**
 * Five defects a bug hunt reproduced in this file's subject, every one of them
 * silently wrong output rather than a failure anybody would see.
 */
describe("a cell the engine must not quietly rewrite", () => {
  it("refuses an impossible date instead of rolling it into a real one", () => {
    // Date.UTC NORMALISES rather than rejecting, so 29 February in a common
    // year became 1 March and 31 April became 1 May — dates a reader believes,
    // printed across every merged slide. The governing rule in recordset.ts is
    // the opposite: a deck that draws perfectly and is two months wrong is
    // worse than one showing the cell untouched.
    for (const cell of ["29/02/2025", "31/04/2026", "31/02/2026", "13/13/2026", "00/00/2026"]) {
      expect(applyFormat(cell, "date:d MMM yyyy"), cell).toBe(cell);
    }
  });

  it("still accepts every date that is real, including a leap day", () => {
    // The refusal must not be a blanket one: 29 February 2024 exists.
    expect(applyFormat("29/2/24", "date:dd-MM-yyyy")).toBe("29-02-2024");
    expect(applyFormat("31/12/2026", "date:dd-MM-yyyy")).toBe("31-12-2026");
    expect(applyFormat("01/01/2026", "date:dd-MM-yyyy")).toBe("01-01-2026");
  });

  it("reads a named or ISO date in the same zone it prints it", () => {
    // `new Date("1 Mar 2026")` parses in the LOCAL zone while formatDate reads
    // UTC fields, so east of UTC every such date came out a day early — in
    // Europe/Copenhagen, this project's own locale, "1 Mar 2026" rendered as
    // "28 Feb 2026". CI runs in UTC, and the only date assertion here used the
    // date-only ISO form the spec parses as UTC, so nothing caught it.
    const tz = process.env.TZ;
    try {
      process.env.TZ = "Europe/Copenhagen";
      expect(applyFormat("1 Mar 2026", "date:dd MMM yyyy")).toBe("01 Mar 2026");
      expect(applyFormat("3 March 2026", "date:dd-MM-yyyy")).toBe("03-03-2026");
      expect(applyFormat("2026-03-01", "date:dd MMM yyyy")).toBe("01 Mar 2026");
    } finally {
      process.env.TZ = tz;
    }
  });

  it("parses a number with more than one thousands group", () => {
    // `replace` without /g changed only the FIRST separator, so "1,234,567"
    // became "1234,567" and then NaN — while detectType still called the column
    // a number, so half of it rendered formatted and half rendered raw.
    expect(numericValue("1,234,567")).toBe(1234567);
    expect(numericValue("1.234.567")).toBe(1234567);
    expect(numericValue("1,234,567.89")).toBe(1234567.89);
    // And the genuinely ambiguous single group still reads as a decimal.
    expect(numericValue("1,5")).toBe(1.5);
  });

  it("returns the cell unchanged when a format asks for impossible decimals", () => {
    // toFixed throws outside 0..100, and `number:-1` is natural to write —
    // Excel's ROUND takes negative digits. Thrown, it killed the whole merge
    // with a message naming neither slide nor placeholder, on a path whose own
    // contract is to return the value unchanged.
    for (const spec of ["number:-1", "number:101", "number:200", "number:abc"]) {
      expect(() => applyFormat("1234.5", spec), spec).not.toThrow();
    }
    expect(applyFormat("1234.5", "number:-1")).toBe("1234.5");
    expect(applyFormat("1234.5", "number:200")).toBe("1234.5");
  });

  it("names a wide table of repeated headers without freezing the pane", () => {
    /**
     * The de-duplication scanned an ARRAY of the names taken so far, and
     * restarted its suffix at 2 for every column — so a header row repeating
     * one name cost O(columns³). This runs inside the `input` handler on the
     * paste box, synchronously, on the UI thread, on every keystroke: 2000
     * repeated columns froze the tab for 2.8 seconds per key, and 4000 for
     * about 22.
     *
     * A pasted range is whatever the user has, and Excel allows 16384 columns.
     * The same shape as the placeholder pattern's quadratic backtracking that
     * `SECURITY.md` records — the impact is a frozen tab either way, and the
     * pane's own call timeouts cannot save it because the work is synchronous.
     *
     * The budget is three orders of magnitude above what the fix costs (1 ms)
     * and an order below what the defect cost, so it is a gate rather than a
     * stopwatch.
     */
    const width = 4000;
    const header = Array.from({ length: width }, () => "Amount");
    const started = Date.now();
    const rs = toRecordSet([header, header.map(() => "1")]);
    expect(Date.now() - started, "naming a wide table is not cubic").toBeLessThan(3000);
    // And it still names them the way it always did.
    expect(rs.columns.length).toBe(width);
    expect(rs.columns.slice(0, 4).map((c) => c.name)).toEqual(["Amount", "Amount 2", "Amount 3", "Amount 4"]);
    expect(rs.columns[width - 1]?.name).toBe(`Amount ${width}`);
    expect(new Set(rs.columns.map((c) => c.name)).size, "every name distinct").toBe(width);
  });

  it("trims a header cell, so a padded one still binds", () => {
    // A pasted range brings whatever whitespace its export gave it. Without the
    // trim the column is declared as `" Name "`, `{{Name}}` binds to nothing,
    // and every slide keeps its placeholder — with the column still listed in
    // the pane under a name that looks right.
    const rs = toRecordSet([
      [" Name ", "\tTown\t"],
      ["Ada", "Aarhus"],
    ]);
    expect(rs.columns.map((c) => c.name)).toEqual(["Name", "Town"]);
    expect(rs.rows[0]?.Name).toBe("Ada");
  });

  it("trims the header on the DEDUP side too, not only where the name is built", () => {
    /**
     * The same `.trim()` appears twice: once building the set of names the
     * sheet DECLARES, and once building each column's name. Only the second had
     * a test, so dropping the first left the whole suite green.
     *
     * It decides whether an INVENTED name may take a name a later column really
     * owns. With `declared` holding `" Column 1 "`, an empty first header
     * invents `Column 1` for itself and the real one is pushed to `Column 1 2`
     * — so a template's `{{Column 1}}` binds to the empty column and prints
     * nothing on every slide, which is the defect the set was added to prevent,
     * reached by a padded header.
     *
     * Found by `scripts/mutate-core.mjs`.
     */
    const padded = toRecordSet([
      ["", " Column 1 "],
      ["a", "b"],
    ]);
    expect(padded.columns.map((c) => c.name)).toEqual(["Column 1 2", "Column 1"]);
    // The row's data goes with the name, which is the half that matters.
    expect(padded.rows[0]?.["Column 1"]).toBe("b");
    // And a header with no padding answers the same way, which is what makes
    // the pair a statement about the trim rather than about this input.
    expect(
      toRecordSet([
        ["", "Column 1"],
        ["a", "b"],
      ]).columns.map((c) => c.name),
    ).toEqual(["Column 1 2", "Column 1"]);
  });

  it("keeps the first row as data when told there is no header", () => {
    // `header: false` is a public option — `toRecordSet` is exported from
    // `src/index.ts` — and it appears nowhere in this repo, so nothing held it.
    // `opts.header ?? true` written as `||` would silently ignore it, which is
    // the shape a tidy-up produces.
    const rs = toRecordSet(
      [
        ["Ada", "Aarhus"],
        ["Grace", "Vejle"],
      ],
      { header: false },
    );
    expect(rs.columns.map((c) => c.name)).toEqual(["Column 1", "Column 2"]);
    expect(rs.rows.length, "the first row is data, not names").toBe(2);
    expect(rs.rows[0]?.["Column 1"]).toBe("Ada");
  });

  it("never lets an invented column name take one a real header owns", () => {
    // Counting forward alone let a made-up name STEAL a later column's: a
    // template's {{Name 2}} bound to the second "Name" and printed the wrong
    // column on every slide, with no warning anywhere.
    const a = toRecordSet([
      ["Name", "Name", "Name 2"],
      ["a", "b", "c"],
    ]);
    expect(a.columns.map((c) => c.name)).toEqual(["Name", "Name 3", "Name 2"]);
    expect(a.rows[0]?.["Name 2"]).toBe("c");

    const b = toRecordSet([
      ["", "Column 1"],
      ["x", "y"],
    ]);
    expect(b.columns.map((c) => c.name)).toEqual(["Column 1 2", "Column 1"]);
    expect(b.rows[0]?.["Column 1"]).toBe("y");
  });
});

describe("which delimiter a paste uses", () => {
  it("reads a quoted FIRST header cell containing a newline", () => {
    // The sniff sampled `src.slice(0, src.indexOf("\n") + 1)`, and that first
    // newline is INSIDE the first header cell here — so it never reached the
    // tab and the whole table parsed as one column. Excel writes a quoted
    // newline whenever a cell holds a line break, and it is legal CSV.
    expect(parseDelimited('"a\nb"\tc\nx\ty')).toEqual([
      ["a\nb", "c"],
      ["x", "y"],
    ]);
  });

  it("still reads a comma paste whose first cell wraps", () => {
    expect(parseDelimited('"a\nb",c\nx,y')).toEqual([
      ["a\nb", "c"],
      ["x", "y"],
    ]);
  });

  it("is not fooled by a doubled quote inside the first cell", () => {
    // `""` is an escaped quote, not the end of the quoted run. Getting that
    // wrong flips the state and the newline after it looks unquoted again.
    expect(parseDelimited('"say ""hi""\nagain"\tc\nx\ty')).toEqual([
      ['say "hi"\nagain', "c"],
      ["x", "y"],
    ]);
  });

  it("reads a tab paste with no quotes at all, as before", () => {
    expect(parseDelimited("First\tLast\nAda\tLovelace")).toEqual([
      ["First", "Last"],
      ["Ada", "Lovelace"],
    ]);
  });

  it("reads a one-line paste, where there is no newline to find", () => {
    expect(parseDelimited("First\tLast")).toEqual([["First", "Last"]]);
    expect(parseDelimited("First,Last")).toEqual([["First", "Last"]]);
  });

  it("reads a quote in the MIDDLE of a header cell as a quote, not as an opening one", () => {
    // The other half of the same rule, and the sniff had it backwards. A quote
    // only opens a quoted cell at the START of a cell — `Size 6" pipe` is a
    // cell containing an inch mark, and RFC 4180 and the parser both read it
    // that way. The sniff toggled on any quote, so it stayed "inside a quote"
    // for the rest of the paste, ran past the first row, found a tab further
    // down and read this comma-separated table as tab-separated: one column,
    // every placeholder unmatched, and nothing on screen saying why.
    //
    // The tab in the body is what makes the paste ambiguous to a scanner that
    // reaches it; there is no tab in the first row, so the answer is a comma.
    expect(parseDelimited('Name,Size 6" pipe\nAda,x\ty')).toEqual([
      ["Name", 'Size 6" pipe'],
      ["Ada", "x\ty"],
    ]);
  });

  it("still opens a quoted cell that follows the OTHER delimiter", () => {
    // The sniff cannot know which delimiter it is about to choose, so it treats
    // both as cell boundaries. A tab-separated paste whose second cell is
    // quoted has to keep working, and so does the comma equivalent.
    expect(parseDelimited('a\t"b\nc"\td\nx\ty\tz')).toEqual([
      ["a", "b\nc", "d"],
      ["x", "y", "z"],
    ]);
    expect(parseDelimited('a,"b\nc",d\nx,y,z')).toEqual([
      ["a", "b\nc", "d"],
      ["x", "y", "z"],
    ]);
  });

  it("reads a semicolon-separated paste, which is what European Excel writes", () => {
    // On a machine whose locale uses the comma as a decimal point — Danish,
    // German, French — Excel's CSV separator is `;`. Before this, such a paste
    // was ONE column named "First;Last", every placeholder unmatched.
    expect(parseDelimited("First;Last\nAda;Lovelace")).toEqual([
      ["First", "Last"],
      ["Ada", "Lovelace"],
    ]);
  });

  it("prefers a comma to a semicolon, so nothing that parses today moves", () => {
    // The whole safety argument for adding the semicolon at all: it is only
    // reached when the first row holds neither a tab nor a comma, so the new
    // branch is reachable only where the old rule already gave one column.
    //
    // `1,5` is a European decimal and `;` is its separator — but this header
    // has a comma, so the comma wins and the row splits on it. That is what
    // this paste did before and has to keep doing.
    expect(parseDelimited("a,b\nx;y,z")).toEqual([
      ["a", "b"],
      ["x;y", "z"],
    ]);
    // And a tab still beats both.
    expect(parseDelimited("a\tb\nx;y\tz,w")).toEqual([
      ["a", "b"],
      ["x;y", "z,w"],
    ]);
  });

  it("opens a quoted cell that follows a SEMICOLON too", () => {
    // The sniff scanner and the delimiter list are one fact: a candidate that
    // can be chosen has to be a cell boundary the scanner honours, or a quoted
    // cell after it leaves the scanner inside a quote for the rest of the
    // paste. That is the defect two tests above, in a third delimiter.
    expect(parseDelimited('a;"b\nc";d\nx;y;z')).toEqual([
      ["a", "b\nc", "d"],
      ["x", "y", "z"],
    ]);
  });

  it("sniffs a fixed number of rows, and the number is what decides the answer", () => {
    /**
     * `chooseDelimiter` scores a SAMPLE, and the sample size is not a detail —
     * it decides the answer. A candidate must split every sampled row into the
     * same number of cells, so one ragged row inside the window disqualifies
     * the real delimiter and the whole paste falls back to the comma, arriving
     * as ONE column whose name holds the tabs. That is the failure the file's
     * own comment describes: "a merge button that does nothing".
     *
     * A ragged row is left alone deliberately — it may be a ragged
     * tab-separated paste or a genuine one-column paste whose text holds tabs,
     * and nothing in the text says which. What must not drift is WHERE the
     * window ends, and nothing held that: moving it by one row moves the
     * boundary between these two cases and no test noticed.
     */
    const good = (n: number) => Array.from({ length: n }, (_, i) => `A${i}\tB${i}`);
    // The sample is ten LINES, header included — so the header plus nine good
    // rows fills it and the tenth data row is the first one outside.
    const outside = ["Name\tTown", ...good(9), "ragged\tx\ty"].join("\n");
    expect(parseDelimited(outside)[0], "the tab survives a raggedness the sniff never saw").toEqual(["Name", "Town"]);
    // One row earlier is inside it, and there the sniff refuses the tab — the
    // documented trade rather than a defect.
    const inside = ["Name\tTown", ...good(8), "ragged\tx\ty"].join("\n");
    expect(parseDelimited(inside)[0], "and inside the sample it refuses").toEqual(["Name\tTown"]);
  });

  it("keeps a one-column header that happens to hold a semicolon", () => {
    // The semicolon splits the FIRST row and no other, so it is not a
    // separator and never qualifies. A rule that looked only at the header
    // split this — which is why this rule does not look only at the header.
    expect(parseDelimited("Notes; extra\nfine\nalso")).toEqual([["Notes; extra"], ["fine"], ["also"]]);
  });

  it("reads a Danish sheet whose header ALSO holds a comma", () => {
    // The case that killed the first rule. "Beløb, EUR" puts a comma in the
    // header, so a header-only sniff chose the comma and split `Ada;1,5;…`
    // into `Ada;1` and `5;…`. Decimal commas are everywhere in exactly the data
    // that uses semicolons, so that rule met one at the first opportunity.
    //
    // Three consistent columns on the semicolon against two on the comma, so
    // the semicolon wins on count.
    expect(parseDelimited("Navn;Beløb, EUR;Dato\nAda;1,5;2026-01-01\nGrace;2,25;2026-02-01")).toEqual([
      ["Navn", "Beløb, EUR", "Dato"],
      ["Ada", "1,5", "2026-01-01"],
      ["Grace", "2,25", "2026-02-01"],
    ]);
  });

  it("gives a genuinely ambiguous paste to the comma, and that is a decision", () => {
    // Two columns on the semicolon, two on the comma, both consistent: nothing
    // in the text says which. The tie goes to the commoner file worldwide.
    // Asserted so the answer is a decision somebody can find and change rather
    // than an accident of iteration order.
    expect(parseDelimited("Navn;Beløb, EUR\nAda;1,5")).toEqual([
      ["Navn;Beløb", " EUR"],
      ["Ada;1", "5"],
    ]);
  });

  it("takes the delimiter that gives MORE columns, not the first one that fits", () => {
    // The step that does the work. Both split every row consistently here —
    // the comma into two, the semicolon into three — and more columns is the
    // reading that used every separator in the row.
    expect(parseDelimited("a;b, c;d\nx;y, z;w")).toEqual([
      ["a", "b, c", "d"],
      ["x", "y, z", "w"],
    ]);
  });

  it("is not thrown off by a BLANK LINE in the middle of the paste", () => {
    // A spacer row is ordinary in a copied Excel range, and it cost the whole
    // table. An empty line splits to one empty cell on every candidate, so it
    // could never match a header of two, step 2 rejected the tab, and the comma
    // fallback returned ONE column whose NAME held the tabs —
    // `Name\tRegion\tRev` — with a merge button that does nothing. The same
    // shape #162 fixed for the semicolon, reached a different way.
    expect(parseDelimited("Name\tRegion\n\nAda\tNorth\nGrace\tSouth")).toEqual([
      ["Name", "Region"],
      [""],
      ["Ada", "North"],
      ["Grace", "South"],
    ]);
    const rs = toRecordSet(parseDelimited("Name\tRegion\n\nAda\tNorth\nGrace\tSouth"));
    expect(rs.columns.map((c) => c.name)).toEqual(["Name", "Region"]);
    // The blank line is not a record either — it is a line, not a row.
    expect(rs.rows).toEqual([
      { Name: "Ada", Region: "North" },
      { Name: "Grace", Region: "South" },
    ]);
  });

  it("still refuses a separator that only splits the header, blank line or not", () => {
    // The guard the fix above must not weaken: `Notes; extra` over rows with no
    // semicolon is ONE column, and skipping blank lines does not make the
    // semicolon consistent.
    expect(toRecordSet(parseDelimited("Notes; extra\n\none\ntwo")).columns.map((c) => c.name)).toEqual([
      "Notes; extra",
    ]);
  });

  it("still lets the caller name the delimiter outright", () => {
    // The pane passes "\t" explicitly for a pasted Excel range; the sniff is
    // only for when nobody said.
    expect(parseDelimited('"a\nb",c', "\t")).toEqual([["a\nb,c"]]);
  });
});

describe("a date that is not a real day", () => {
  /**
   * The engine's governing rule, stated in `recordset.ts` and again on
   * `utcDate`: a merged deck that draws perfectly and is a month wrong is worse
   * than one showing the cell untouched. `utcDate` exists to enforce it, by
   * refusing anything that did not survive a round trip.
   *
   * It never fired on the string path, and could not have. `new Date` had
   * already NORMALISED — 29 February 2026 into 1 March — so the components read
   * back off the result were valid ones that round-tripped perfectly. The guard
   * was reached only with numbers something else had made correct.
   *
   * Every accepting form here comes from a spreadsheet or a typed cell, which
   * is where a 29 February in a common year comes from in the first place.
   */
  it("refuses a day the month does not have, in ISO", () => {
    expect(parseDate("2026-02-29"), "29 Feb in a common year").toBeUndefined();
    expect(parseDate("2026-04-31"), "31 April").toBeUndefined();
    expect(parseDate("2026-06-31"), "31 June").toBeUndefined();
  });

  it("refuses a day the month does not have, written with the month's name", () => {
    // These are worse than the ISO ones: `new Date("31 Feb 2026")` answers
    // 3 MARCH, so the month moved as well as the day.
    expect(parseDate("31 Feb 2026")).toBeUndefined();
    expect(parseDate("30 February 2026")).toBeUndefined();
    expect(parseDate("31 Apr 2026")).toBeUndefined();
  });

  it("still accepts a real leap day", () => {
    // Not a guard against the bug — this passed before the fix too. It guards
    // the FIX from being too broad: refusing every 29 February would be a
    // cheap way to make the tests above green and would be worse than the bug.
    expect(parseDate("2024-02-29")?.toISOString().slice(0, 10)).toBe("2024-02-29");
    expect(parseDate("29 Feb 2024")?.toISOString().slice(0, 10)).toBe("2024-02-29");
  });

  it("reads the components from the STRING, not from a parsed Date", () => {
    /**
     * The ISO form with a time is parsed by `new Date` in the host's LOCAL
     * zone, so east of UTC reading the components back gives the previous day.
     * The suite runs in UTC and cannot see that, which is how it shipped once
     * already — so this asserts the shape that makes it impossible rather than
     * the symptom, which no assertion here could reach.
     *
     * Also not a guard against the rollover: it passed before the fix, in this
     * timezone. It is here because the fix moved this path, and the thing it
     * protects is invisible to a suite running in UTC.
     */
    expect(parseDate("2026-03-01T10:00")?.toISOString().slice(0, 10)).toBe("2026-03-01");
    expect(parseDate("2026-03-01 23:59")?.toISOString().slice(0, 10)).toBe("2026-03-01");
  });

  it("returns the raw cell to the slide when it refuses", () => {
    // The point of refusing: the author sees what they typed and can fix it,
    // instead of 240 slides carrying a date nobody wrote.
    expect(applyFormat("2026-02-29", "date")).toBe("2026-02-29");
    expect(applyFormat("31 Feb 2026", "date:d MMM yyyy")).toBe("31 Feb 2026");
  });
});

describe("month names the date gate already admits", () => {
  /**
   * `NAMED_DATE` allows `ÆØÅ`, which is somebody having made room for Danish
   * deliberately. Resolving the name through `new Date` then handled Danish BY
   * ACCIDENT — that parser matches an English three-letter prefix, so `marts`
   * and `januar` worked and `maj` and `desember` did not.
   *
   * The symptom is the sharpest form of the rule this engine is built on: one
   * column, one format spec, two renderings. `detectType` still called it a
   * date, so half the slides carried `01-03-2026` and half carried
   * `3 maj 2026`, with nothing saying why.
   */
  const DANISH: [string, string][] = [
    ["1 januar 2026", "2026-01-01"],
    ["2 februar 2026", "2026-02-02"],
    ["3 marts 2026", "2026-03-03"],
    ["4 april 2026", "2026-04-04"],
    ["5 maj 2026", "2026-05-05"],
    ["6 juni 2026", "2026-06-06"],
    ["7 juli 2026", "2026-07-07"],
    ["8 august 2026", "2026-08-08"],
    ["9 september 2026", "2026-09-09"],
    ["10 oktober 2026", "2026-10-10"],
    ["11 november 2026", "2026-11-11"],
    ["12 december 2026", "2026-12-12"],
  ];

  it.each(DANISH)("reads the Danish %s", (raw, iso) => {
    expect(parseDate(raw)?.toISOString().slice(0, 10)).toBe(iso);
  });

  it("reads the Norwegian and Swedish spellings that differ", () => {
    // The whole reason a table beats a prefix rule: these are the words an
    // English-only answer gets wrong, not the ones it gets right by luck.
    expect(parseDate("1 mars 2026")?.toISOString().slice(0, 10)).toBe("2026-03-01");
    expect(parseDate("1 mai 2026")?.toISOString().slice(0, 10)).toBe("2026-05-01");
    expect(parseDate("1 desember 2026")?.toISOString().slice(0, 10)).toBe("2026-12-01");
    expect(parseDate("1 januari 2026")?.toISOString().slice(0, 10)).toBe("2026-01-01");
    expect(parseDate("1 augusti 2026")?.toISOString().slice(0, 10)).toBe("2026-08-01");
  });

  it("reads the abbreviations a spreadsheet writes", () => {
    expect(parseDate("1 okt 2026")?.toISOString().slice(0, 10)).toBe("2026-10-01");
    expect(parseDate("1 des 2026")?.toISOString().slice(0, 10)).toBe("2026-12-01");
    expect(parseDate("1 Maj 2026")?.toISOString().slice(0, 10)).toBe("2026-05-01");
  });

  it("does not let a Nordic name through as the wrong month", () => {
    /**
     * The property that makes one table safe for four languages, asserted
     * rather than trusted: no word means a different month in a different one
     * of them. Written as a check over the WHOLE table, so a word added later
     * that clashes fails here instead of quietly renaming a month.
     */
    const languages = [
      "january february march april may june july august september october november december",
      "januar februar marts april maj juni juli august september oktober november december",
      "januar februar mars april mai juni juli august september oktober november desember",
      "januari februari mars april maj juni juli augusti september oktober november december",
      // The ASCII spellings of the languages the removed `new Date` fallback
      // used to reach by accident, and their neighbours. The accented forms —
      // `février`, `août`, `décembre`, `märz`, `março` — ARE in the table,
      // and have their own test below; what stays out is the transliteration
      // somebody's exporter might write instead, which is `notClaimed`.
      "januar februar mars april mai juni juli august september oktober november dezember",
      "januari februari maart april mei juni juli augustus september oktober november december",
      "janvier fevrier mars avril mai juin juillet aout septembre octobre novembre decembre",
      "enero febrero marzo abril mayo junio julio agosto septiembre octubre noviembre diciembre",
      "gennaio febbraio marzo aprile maggio giugno luglio agosto settembre ottobre novembre dicembre",
      "janeiro fevereiro marco abril maio junho julho agosto setembro outubro novembro dezembro",
    ];
    // The spellings this table deliberately does NOT carry, because their real
    // form is accented and the ASCII fallback is somebody else's guess. The
    // real form is read; this list is the guess at it, and it stays refused.
    const notClaimed = new Set(["fevrier", "aout", "decembre", "marco"]);
    const meaning = new Map<string, number>();
    for (const line of languages) {
      line.split(" ").forEach((word, i) => {
        const month = i + 1;
        const already = meaning.get(word);
        expect(already ?? month, `${word} means two different months`).toBe(month);
        meaning.set(word, month);
        // And the engine agrees with the list.
        if (notClaimed.has(word)) {
          expect(parseDate(`1 ${word} 2026`), `1 ${word} 2026`).toBeUndefined();
          return;
        }
        expect(parseDate(`1 ${word} 2026`)?.getUTCMonth(), `1 ${word} 2026`).toBe(i);
      });
    }
    expect(meaning.size).toBeGreaterThan(25);
  });

  it("reads the SHORT month for every language the manual names", () => {
    /**
     * The manual claimed ten languages "in full or in the three-letter form a
     * spreadsheet writes", and only the English and Scandinavian abbreviations
     * were in the table. So a German column of `1 dez 2026` typed as a date on
     * SHAPE and then printed as the raw cell — half a deck formatted and half
     * not, with nothing saying why, which is the exact outcome the table was
     * added to prevent, arriving through it rather than around it.
     *
     * Same property as the full names above, asserted the same way: no short
     * form means a different month in a different language. `mar` is March in
     * three of them, `dez` is December in two, `set` is September in two.
     */
    const shortForms = [
      "jan feb mar apr may jun jul aug sep oct nov dec",
      "jan feb mar apr maj jun jul aug sep okt nov dec",
      "jan feb mar apr mai jun jul aug sep okt nov des",
      // German as Excel on Windows writes it. The accented `m\u00e4r` is CLDR's
      // spelling and has its own line below; a German column pasted out of
      // Excel carries `Mrz`, and the table read eleven of its twelve months
      // until that was noticed — which is the half-formatted deck again,
      // missing by one word. The row is here because the list below had eight
      // lines for ten languages and German was never one of them.
      "jan feb mrz apr mai jun jul aug sep okt nov dez",
      // Dutch, Spanish, Italian, Portuguese.
      "jan feb mrt apr mei jun jul aug sep okt nov dez",
      "ene feb mar abr may jun jul ago sep oct nov dic",
      "gen feb mar apr mag giu lug ago set ott nov dic",
      "jan fev mar abr mai jun jul ago set out nov dez",
      // French writes four letters for four of them, and a three-letter
      // truncation is not a form anybody exports.
      "janv f\u00e9vr mars avr mai juin juil ao\u00fbt sept oct nov d\u00e9c",
    ];
    const meaning = new Map<string, number>();
    for (const line of shortForms) {
      line.split(" ").forEach((word, i) => {
        const already = meaning.get(word);
        expect(already ?? i + 1, `${word} means two different months`).toBe(i + 1);
        meaning.set(word, i + 1);
        expect(parseDate(`1 ${word} 2026`)?.getUTCMonth(), `1 ${word} 2026`).toBe(i);
      });
    }
    expect(parseDate("1 m\u00e4r 2026")?.getUTCMonth(), "the German March, CLDR's spelling").toBe(2);
    expect(parseDate("1 Mrz 2026")?.getUTCMonth(), "the German March, Excel's spelling").toBe(2);
    // And a whole column of one language formats, rather than half of it.
    expect(applyFormat("1 dez 2026", "date")).toBe("01-12-2026");
    expect(applyFormat("1 gen 2026", "date")).toBe("01-01-2026");
    expect(applyFormat("1 out 2026", "date")).toBe("01-10-2026");
  });

  it("does not turn an ordinary English phrase into a date", () => {
    // `out` is October in Portuguese and an English word, and `set` and `ago`
    // are the same shape. They are safe because `NAMED_DATE` admits a month
    // word only between a one-or-two-digit number and a two-to-four-digit
    // year — so the phrase has to look exactly like a date before the table is
    // ever asked.
    expect(parseDate("2 out 3")).toBeUndefined();
    expect(parseDate("5 set 9")).toBeUndefined();
    expect(parseDate("2 days ago")).toBeUndefined();
    expect(detectType(["2 out 3", "5 set 9"])).toBe("text");
  });

  it("still refuses a day that month does not have, whatever the language", () => {
    // The table must not become a way round the guard added with it.
    expect(parseDate("31 apr 2026")).toBeUndefined();
    expect(parseDate("30 februar 2026")).toBeUndefined();
    expect(parseDate("32 maj 2026")).toBeUndefined();
  });

  it("keeps whatever the platform used to handle, from the table instead", () => {
    /**
     * There was a `new Date` fallback behind the table, kept on purpose so that
     * languages the table did not list still worked. It is gone — see
     * `monthFromName` — and these still pass, because what it really covered is
     * listed now. More of it is: the prefix rule read `janvier` and `marzo` by
     * luck and answered nothing at all for `enero`, `maart` or `gennaio`.
     */
    expect(parseDate("1 janvier 2026")?.toISOString().slice(0, 10)).toBe("2026-01-01");
    expect(parseDate("1 marzo 2026")?.toISOString().slice(0, 10)).toBe("2026-03-01");
    expect(parseDate("1 enero 2026")?.toISOString().slice(0, 10)).toBe("2026-01-01");
    expect(parseDate("1 gennaio 2026")?.toISOString().slice(0, 10)).toBe("2026-01-01");
    expect(parseDate("1 maart 2026")?.toISOString().slice(0, 10)).toBe("2026-03-01");
    expect(parseDate("1 March 2026")?.toISOString().slice(0, 10)).toBe("2026-03-01");
  });

  it("refuses a word that merely STARTS like a month", () => {
    /**
     * The defect the table was written to prevent, living in the branch behind
     * it. `new Date("1 marketing 2001 00:00:00Z")` is 1 March 2001, because V8
     * matches an English three-letter prefix — so a cell reading
     * `1 marketing 2026` was typed `date`, parsed as March, and printed as
     * `01-03-2026` on every merged slide.
     *
     * Each of these is an ordinary English word with a month's first three
     * letters, and every one of them parsed before this.
     */
    for (const [word, wrongly] of [
      ["marketing", "March"],
      ["janitor", "January"],
      ["novel", "November"],
      ["decision", "December"],
      ["separate", "September"],
      ["octopus", "October"],
      ["junior", "June"],
    ] as const) {
      expect(parseDate(`1 ${word} 2026`), `read as ${wrongly}`).toBeUndefined();
      expect(applyFormat(`1 ${word} 2026`, "date:dd-MM-yyyy"), word).toBe(`1 ${word} 2026`);
    }
    // Not `augustus`, which looks like the same shape and is Dutch for August.
    // A word that starts like a month is not automatically a non-month, which
    // is why the list above is words rather than a rule.
    expect(parseDate("1 augustus 2026")?.getUTCMonth()).toBe(7);
  });

  it("reads the accented spellings, which the character class used to refuse", () => {
    /**
     * `NAMED_DATE` admitted `[A-Za-z\u00c6\u00d8\u00c5\u00e6\u00f8\u00e5]`, so French, German and
     * Portuguese were PARTIAL: nine months of twelve read and three did not.
     * That is one column rendering two ways, which is the defect this whole
     * describe block exists for — and the sharpest version of it, because the
     * three that failed are the ones nobody spells any other way.
     */
    expect(parseDate("1 f\u00e9vrier 2026")?.toISOString().slice(0, 10)).toBe("2026-02-01");
    expect(parseDate("1 ao\u00fbt 2026")?.toISOString().slice(0, 10)).toBe("2026-08-01");
    expect(parseDate("1 d\u00e9cembre 2026")?.toISOString().slice(0, 10)).toBe("2026-12-01");
    expect(parseDate("1 m\u00e4rz 2026")?.toISOString().slice(0, 10)).toBe("2026-03-01");
    expect(parseDate("1 mar\u00e7o 2026")?.toISOString().slice(0, 10)).toBe("2026-03-01");
    // Capitalised, which is how German writes it and how a spreadsheet's own
    // month formatting writes all of them.
    expect(parseDate("1 M\u00e4rz 2026")?.toISOString().slice(0, 10)).toBe("2026-03-01");
  });

  it("reads a name whose accent is a combining mark", () => {
    /**
     * The same five words in NFD, which is what a CSV exported on macOS
     * carries: `\u00e9` is `e` followed by U+0301 rather than one code point. The
     * character class has to admit the mark and the table lookup has to
     * normalise, and neither half is visible in a string that LOOKS identical
     * in a diff.
     */
    const nfd = "1 f\u00e9vrier 2026".normalize("NFD");
    expect(nfd).not.toBe("1 f\u00e9vrier 2026");
    expect(parseDate(nfd)?.toISOString().slice(0, 10)).toBe("2026-02-01");
    expect(parseDate("1 ao\u00fbt 2026".normalize("NFD"))?.toISOString().slice(0, 10)).toBe("2026-08-01");
  });

  it("formats a whole French column the same way, which was the same defect", () => {
    // Before this, the first of these formatted and the other two came out raw.
    const column = ["1 janvier 2026", "1 f\u00e9vrier 2026", "1 ao\u00fbt 2026"];
    expect(column.map((c) => applyFormat(c, "date:dd-MM-yyyy"))).toEqual(["01-01-2026", "01-02-2026", "01-08-2026"]);
  });

  it("formats a whole Danish column the same way, which was the defect", () => {
    // The end-to-end shape: before the table, `maj` fell through to the raw
    // cell while `marts` formatted, in the same column, on the same deck.
    const column = ["1 marts 2026", "3 maj 2026", "1 desember 2026"];
    expect(column.map((c) => applyFormat(c, "date:dd-MM-yyyy"))).toEqual(["01-03-2026", "03-05-2026", "01-12-2026"]);
  });
});

describe("formats that were wrong in small ways", () => {
  it("does not print minus zero", () => {
    // The sign came from the INPUT, not the rounded value, so `-0.4` at no
    // decimal places printed `-0` — a quantity that does not exist, on a
    // slide, from an ordinary cell.
    expect(applyFormat("-0.4", "number")).toBe("0");
    expect(applyFormat("-0.04", "number:1")).toBe("0,0");
    // And still keeps the sign where there is one to keep.
    expect(applyFormat("-1.5", "number")).toBe("-2");
    expect(applyFormat("-0.4", "number:1")).toBe("-0,4");
  });

  it("takes only a count of places as a decimal count", () => {
    // `Number` also reads `1e2`, so `number:1e2` asked for a hundred decimal
    // places and got them: a legal `toFixed` call producing a number no slide
    // has room for. A spec that is not a count is not a count, and the cell is
    // returned as it stands like every other unreadable format.
    expect(applyFormat("1234.5", "number:1e2")).toBe("1234.5");
    expect(applyFormat("1234.5", "number:0x2")).toBe("1234.5");
    expect(applyFormat("1234.5", "number:-1")).toBe("1234.5");
    expect(applyFormat("1234.5", "number:101")).toBe("1234.5");
    // The forms that ARE a count still work, surrounding space included.
    expect(applyFormat("1234.5", "number:2")).toBe("1 234,50");
    expect(applyFormat("1234.5", "number: 2 ")).toBe("1 234,50");
  });

  it("writes a full month name for MMMM", () => {
    // It printed `MarM` — `MMM` replaced and the fourth `M` left standing.
    // Supported rather than refused, since a full month is a thing to want.
    const d = parseDate("2026-03-01");
    expect(d && formatDate(d, "MMMM")).toBe("March");
    expect(d && formatDate(d, "d MMMM yyyy")).toBe("1 March 2026");
    // And the shorter tokens still mean what they meant.
    expect(d && formatDate(d, "MMM")).toBe("Mar");
    expect(d && formatDate(d, "yyyy-MM-dd")).toBe("2026-03-01");
  });

  it("leaves the letters of a literal alone", () => {
    // The `\b` around the single `d` is load-bearing: without it a pattern of
    // `Ends d` prints "En1s 1". Longest-token-first is the other half — swap
    // `yyyy` and `yy` and the year becomes "2626".
    const d = parseDate("2026-03-01");
    expect(d && formatDate(d, "Ends d")).toBe("Ends 1");
    expect(d && formatDate(d, "yyyy")).toBe("2026");
  });
});

describe("one definition of a number", () => {
  /**
   * `detectType` asked a regex; `numericValue` asked `Number()`, which is a far
   * wider gate. A column of product codes reading `0x10` was called text and
   * converted as sixteen, and the two answers met on the same slide.
   *
   * The invariant is the assertion, not the values: whatever the engine is
   * willing to CONVERT it must also be willing to CALL a number. That is what
   * stays true when either side is edited next.
   */
  const NOT_NUMBERS = ["0x10", "0b11", "0o17", "1e3", "1E3", ".5", "+7"];

  it.each(NOT_NUMBERS)("refuses %s in both places at once", (raw) => {
    expect(looksLikeNumber(raw)).toBe(false);
    expect(numericValue(raw)).toBeUndefined();
    expect(detectType([raw])).toBe("text");
  });

  it("leaves a refused cell exactly as the author typed it", () => {
    // A product code silently turned into 16 across a merged deck reads as
    // deliberate, which is what makes it worse than a visible refusal.
    expect(applyFormat("0x10", "number:0")).toBe("0x10");
    expect(applyFormat("1e3", "number:0")).toBe("1e3");
  });

  it("still reads the numbers a spreadsheet actually writes", () => {
    expect(numericValue("  12  ")).toBe(12);
    expect(numericValue("-7")).toBe(-7);
    expect(numericValue("1234.5")).toBe(1234.5);
    expect(detectType(["12", "-7", "1234.5"])).toBe("number");
  });
});

describe("the day written as an ordinal", () => {
  /**
   * `1. marts 2026` is the ordinary Danish long form. `NAMED_DATE` required
   * exactly one separator character, so it admitted `1 marts 2026` and
   * `1.marts 2026` and refused the form people type — leaving the month-name
   * table above reachable mainly by spellings nobody writes.
   *
   * Widening the gate ALONE made it briefly worse: `parseDate` carried a
   * private copy of the pattern, so a column typed as `date` and then rendered
   * raw. That is the two-renderings failure again, entered from the other
   * side. One exported regex now answers for both.
   */
  const FORMS = ["1. marts 2026", "1 marts 2026", "1.marts 2026", "1. mar 2026", "1. March 2026"];

  it.each(FORMS)("reads %s as the first of March", (raw) => {
    expect(looksLikeDate(raw)).toBe(true);
    expect(parseDate(raw)?.toISOString().slice(0, 10)).toBe("2026-03-01");
    expect(detectType([raw])).toBe("date");
    expect(applyFormat(raw, "date:d MMM yyyy")).toBe("1 Mar 2026");
  });

  it("keeps the gate and the parser answering together", () => {
    // Not a universal law, and saying so is the point: a well-formed date can
    // still be an impossible one, and `31 Feb 2026` is meant to pass the gate
    // and fail the parse. They must agree for every value that is BOTH well
    // formed and real, which is where the drift above actually hurt.
    for (const raw of [...FORMS, "1 January 2026", "2026-03-01", "1. januar 26"]) {
      expect(looksLikeDate(raw)).toBe(true);
      expect(parseDate(raw)).toBeDefined();
    }
  });

  it("still refuses what was always ambiguous", () => {
    // Punctuation was loosened. Ambiguity was not.
    expect(looksLikeDate("03/01/2026")).toBe(false);
    expect(detectType(["03/01/2026"])).toBe("text");
  });
});

describe("the number gate and the number parser cannot disagree", () => {
  /**
   * `detectType` asks `looksLikeNumber`; `applyFormat` asks `numericValue`. A
   * form one accepts and the other cannot read is a column typed as a number
   * that renders raw — the failure this pair has now produced twice, once
   * through `Number()` and once through the grouping pattern.
   *
   * `1,234,5` is nobody's locale, but the gate admitted it and the parser
   * returned undefined, so it is the same defect however unlikely the cell.
   * The pattern now captures the grouping separator, makes the later groups
   * repeat THAT one, and forbids a decimal part from reusing it.
   */
  const CORPUS = (() => {
    // Every character that can separate two digits here, which since
    // 2026-09-01 is four spaces rather than one: the gate admits the no-break,
    // narrow no-break and thin spaces Excel groups with.
    //
    // Widening this list detects NOTHING TODAY, and that is worth saying out
    // loud rather than leaving a reader to assume the opposite. The property
    // below is a tautology — `looksLikeNumber` is defined as `numericValue`,
    // which is the right end state — so the corpus is a record of the input
    // space this pair is meant to cover, kept in step with the pattern so that
    // it is already right if the two are ever separated again.
    const seps = ["", ",", ".", " ", "\u00a0", "\u202f", "\u2009"];
    const groups = ["", "0", "5", "12", "234", "1234"];
    const out = new Set<string>();
    for (const a of groups)
      for (const s1 of seps)
        for (const b of groups)
          for (const s2 of seps)
            for (const c of groups) {
              const v = a + s1 + b + s2 + c;
              if (v === "" || !/\d/.test(v)) continue;
              out.add(v);
              out.add("-" + v);
            }
    return [...out];
  })();

  it("agrees about every arrangement of digits and separators", () => {
    // Swept rather than listed, because the shapes that broke it were the ones
    // nobody thought to write down. Reverting the pattern to its previous form
    // puts 64 of these back.
    const violations = CORPUS.filter((v) => looksLikeNumber(v) !== (numericValue(v) !== undefined));
    expect(violations, "a form one of them accepts and the other cannot read").toEqual([]);
    expect(CORPUS.length, "the sweep stopped covering anything").toBeGreaterThan(9000);
  });

  it("refuses a separator used for both grouping and the decimal", () => {
    for (const v of ["1,234,5", "1.234.5", "1 234 5"]) {
      expect(looksLikeNumber(v), v).toBe(false);
      expect(numericValue(v), v).toBeUndefined();
    }
  });

  it("reads a lone three-digit group the same way for both separators", () => {
    /**
     * The comment beside the dot branch claimed a single `1.234` stayed a
     * decimal. It never did — `+` is one or more, so one group has always been
     * read as grouping, which is also what the comma branch does with `1,500`.
     * Pinned here because the two readings differ by a factor of a thousand and
     * the only thing worse than picking one is not knowing which was picked.
     */
    expect(numericValue("1.234")).toBe(1234);
    expect(numericValue("1,234")).toBe(1234);
    expect(numericValue("1 234")).toBe(1234);
    // Two digits after the separator is not a group, and stays a decimal.
    expect(numericValue("1.23")).toBe(1.23);
    expect(numericValue("1,23")).toBe(1.23);
  });
});

describe("a number too large for a double", () => {
  /**
   * The one input class where making `looksLikeNumber` ask `numericValue`
   * CHANGED an answer rather than preserving it.
   *
   * The pattern matches any run of digits, and `Number()` gives up above about
   * 1.8e308. So a 310-digit cell used to be typed `number` by the detector and
   * refused by the converter — the exact disagreement that change was made to
   * end, in an input class no test covered. It is typed `text` now, and both
   * halves say so.
   *
   * Text is the honest answer rather than the ideal one. A spreadsheet would
   * call it a number; this engine cannot hold it as one, and claiming it is a
   * number and then failing to format it is the worse of the two — the same
   * reasoning that refuses an ambiguous date rather than guessing it.
   */
  const HUGE = "9".repeat(310);

  it("is refused by both halves rather than one", () => {
    expect(numericValue(HUGE)).toBeUndefined();
    expect(looksLikeNumber(HUGE)).toBe(false);
    expect(detectType([HUGE])).toBe("text");
  });

  it("and the largest one that DOES fit still counts", () => {
    // The other half: this refuses values, not long strings of digits.
    expect(looksLikeNumber("1.7976931348623157")).toBe(true);
    expect(numericValue("9".repeat(15))).toBe(999999999999999);
    expect(detectType(["9".repeat(15)])).toBe("number");
  });
});

describe("a data cell cannot reach Object's prototype", () => {
  /**
   * From the security sweep of 2026-08-30. `monthFromName` looks a month word
   * up in a table keyed by the cell's own text, and `dateShape`'s `NAMED_DATE`
   * takes ANY word of three letters or more — so `1 constructor 2026` passes
   * the gate and reaches the lookup.
   *
   * `Object.freeze` does not remove the prototype chain. Unguarded, the table
   * answers the `Object` function for `constructor` and `Object.prototype` for
   * `__proto__`, and `known !== undefined` accepts either as a month.
   *
   * The OUTCOME was already correct and only by luck — a function reaches
   * `Date.UTC`, the arithmetic is NaN, the date is invalid, and the rule that
   * an unreadable cell is printed as it stands catches it.
   *
   * **So these are CHARACTERISATION tests and not a proof of the fix, and that
   * was checked rather than assumed: with the guard reverted they still pass,
   * all 92 of them.** This repo's rule is that a test which passes against the
   * unfixed file is decoration, and calling these a guard would be exactly
   * that. What they are for is the day the luck runs out — a `dateFrom` that
   * tolerates a non-number, or a table whose prototype carries a numeric
   * property — at which point they go red and the guard beside them is why
   * they do not need to.
   */
  for (const word of ["constructor", "__proto__", "hasownproperty", "tostring", "valueof"]) {
    it(`leaves "1 ${word} 2026" alone`, () => {
      expect(applyFormat(`1 ${word} 2026`, "date")).toBe(`1 ${word} 2026`);
      expect(applyFormat(`1 ${word} 2026`, "date:d MMM yyyy")).toBe(`1 ${word} 2026`);
    });
  }

  it("still reads the months it is supposed to", () => {
    // The other half: a guard that refused everything would pass the tests
    // above and break the feature.
    expect(applyFormat("1 january 2026", "date:d MMM yyyy")).toBe("1 Jan 2026");
    expect(applyFormat("1 januar 2026", "date:d MMM yyyy")).toBe("1 Jan 2026");
    // And an abbreviation, which is in the table like everything else — there
    // is no `new Date` branch behind it any more.
    expect(applyFormat("1 Mar 2026", "date:d MMM yyyy")).toBe("1 Mar 2026");
  });
});

describe("dates a spreadsheet actually writes", () => {
  it("reads a two-digit year in the window Excel uses", () => {
    /**
     * Excel reads 00-29 as the 2000s and 30-99 as the 1900s, and `dd/mm/yy` is
     * an ordinary cell format there. This added 2000 to everything, so
     * `15/06/85` merged as 2085 — a birth date or a contract history a century
     * out on every slide, silently, with the column still typed a clean date.
     * The tests only ever used 24 and 26, which are on the right side of the
     * cutoff under either rule.
     */
    expect(applyFormat("15/06/85", "date:dd-MM-yyyy")).toBe("15-06-1985");
    expect(applyFormat("15 juni 85", "date:d MMMM yyyy")).toBe("15 June 1985");
    // The boundary, from both sides.
    expect(applyFormat("01/01/29", "date:yyyy")).toBe("2029");
    expect(applyFormat("01/01/30", "date:yyyy")).toBe("1930");
    // A four-digit year is untouched.
    expect(applyFormat("15/06/1985", "date:yyyy")).toBe("1985");
  });

  it("refuses an ISO date with something other than a time after it", () => {
    /**
     * The tail was `.*`, so a PERIOD — an ordinary way to write a range — was
     * typed as a date and formatted as its first half, with the rest discarded
     * and nothing to say so.
     */
    expect(applyFormat("2026-03-01 - 2026-03-31", "date:d MMM yyyy")).toBe("2026-03-01 - 2026-03-31");
    expect(applyFormat("2026-03-01 (provisional)", "date")).toBe("2026-03-01 (provisional)");
    expect(looksLikeDate("2026-03-01 - 2026-03-31")).toBe(false);
    // A real time still reads, which is what the tail was for.
    expect(applyFormat("2026-03-01T10:00", "date:d MMM yyyy")).toBe("1 Mar 2026");
    expect(applyFormat("2026-03-01 10:00:00", "date:d MMM yyyy")).toBe("1 Mar 2026");
    expect(applyFormat("2026-03-01", "date:d MMM yyyy")).toBe("1 Mar 2026");
  });

  it("prints text in a date pattern as written, including words holding dd", () => {
    // The manual promises exactly this. Only the single `d` was bounded, and
    // `dd` sits inside add, odd, middle, wedding and address.
    expect(applyFormat("2026-03-01", "date:Odds: d MMM yyyy")).toBe("Odds: 1 Mar 2026");
    expect(applyFormat("2026-03-01", "date:Wedding on d MMMM")).toBe("Wedding on 1 March");
    // And the tokens themselves still work, longest first.
    expect(applyFormat("2026-03-01", "date:dd/MM/yyyy")).toBe("01/03/2026");
    expect(applyFormat("2026-03-01", "date:d MMMM yy")).toBe("1 March 26");
  });

  it("reads a pattern made of nothing but tokens", () => {
    /**
     * The other direction, and the one the bound broke. Written as a lookbehind
     * for "not next to any letter", a pattern of ADJACENT tokens refused to
     * match at all: `yyyyMMdd` — the ISO-basic spelling used for file names,
     * and every token in it one the manual lists — printed itself. A pattern of
     * nothing but tokens printed as written is exactly what the manual says
     * will not happen.
     */
    expect(applyFormat("2026-03-01", "date:yyyyMMdd")).toBe("20260301");
    expect(applyFormat("2026-03-01", "date:ddMMyyyy")).toBe("01032026");
    expect(applyFormat("2026-03-01", "date:yyMMdd")).toBe("260301");
    // A run that is part token and part word is a WORD. `ddTHH` holds letters
    // that are not tokens, so it prints as written — which is the manual's own
    // rule, and the price of it is that `yyyy-MM-ddTHH:mm` does not read as an
    // ISO datetime. Taking the other trade ("starts with a token, tail printed
    // as written") turns `den` into `1en`, and `den` is the ordinary Danish,
    // German and Swedish long-date word.
    expect(applyFormat("2026-03-01", "date:yyyy-MM-ddTHH:mm")).toBe("2026-03-ddTHH:mm");
  });

  it("prints a month name written into the pattern, in any language", () => {
    /**
     * The output month is English, so a template author writes their own — and
     * the manual now shows it inside ONE placeholder rather than splitting the
     * date across three. It works because a run of letters is either all tokens
     * or none: `maj`, `Dezember` and `de` are words, and `d` and `yyyy` beside
     * them are not.
     *
     * These are the four examples the manual prints. A doc that shows an output
     * nobody checked is the same defect as a stale comment.
     */
    expect(applyFormat("2026-03-01", "date:d. maj yyyy")).toBe("1. maj 2026");
    expect(applyFormat("2026-03-01", "date:d. Dezember yyyy")).toBe("1. Dezember 2026");
    expect(applyFormat("2026-03-01", "date:d de mayo de yyyy")).toBe("1 de mayo de 2026");
    expect(applyFormat("2026-03-01", "date:den d. MMMM yyyy")).toBe("den 1. March 2026");
  });

  it("leaves an ACCENTED word that begins with a token's letter alone", () => {
    /**
     * The rule is about words, and matching ASCII letters alone cut a word into
     * runs at its own accent — leaving a one-character run of `d`, which is a
     * whole token. `día d de MMMM` printed `1ía 1 de March` and
     * `Date d'échéance dd/MM/yyyy` printed `Date 1'échéance 01/03/2026`.
     *
     * Spanish, French, Danish and Swedish all write ordinary date words that
     * begin with a token's letter and continue past ASCII, and the manual
     * promises text in a pattern prints as written.
     */
    expect(applyFormat("2026-03-01", "date:día d de MMMM")).toBe("día 1 de March");
    expect(applyFormat("2026-03-01", "date:décembre yyyy")).toBe("décembre 2026");
    expect(applyFormat("2026-03-01", "date:Date d'échéance dd/MM/yyyy")).toBe("Date d'échéance 01/03/2026");
    expect(applyFormat("2026-03-01", "date:dåb d MMM")).toBe("dåb 1 Mar");
    expect(applyFormat("2026-03-01", "date:död d MMM")).toBe("död 1 Mar");
    // U+2019 as well as U+0027: the typographic apostrophe is what Word's
    // autocorrect produces and what a paste from a document carries.
    expect(applyFormat("2026-03-01", "date:Date d\u2019échéance dd/MM/yyyy")).toBe("Date d\u2019échéance 01/03/2026");
    // And nothing that used to tokenize stopped: the tokens are ASCII, so
    // widening what counts as a word can only make a run stop matching.
    expect(applyFormat("2026-03-01", "date:yyyyMMdd")).toBe("20260301");
    expect(applyFormat("2026-03-01", "date:dd-MM-yyyy")).toBe("01-03-2026");
  });

  it("reads a pattern whose own tokens are joined by an apostrophe", () => {
    /**
     * Holding an apostrophe inside a word is what leaves `d'échéance` alone,
     * and it welds `MMM'yy` — `Mar'26`, the commonest abbreviated pattern
     * there is — into a single run that is not a token, so the whole thing
     * printed itself. Word's autocorrect turns the straight apostrophe into
     * U+2019, so `{{Signed|date:MMM’yy}}` reached it by simply being typed in
     * Word.
     *
     * A run that is not all tokens is retried piece by piece, and every piece
     * has to be one: `MMM'yy` reads, `d'échéance` does not, because `échéance`
     * is not a token.
     */
    expect(applyFormat("2026-03-01", "date:MMM'yy")).toBe("Mar'26");
    expect(applyFormat("2026-03-01", "date:MMM\u2019yy")).toBe("Mar\u201926");
    expect(applyFormat("2026-03-01", "date:dd'MM'yyyy")).toBe("01'03'2026");
    expect(applyFormat("2026-03-01", "date:d\u2019MMM")).toBe("1\u2019Mar");
    // One piece that is not a token fails the whole run, which is what keeps
    // the accented-word case above working.
    expect(applyFormat("2026-03-01", "date:Date d'échéance MMM'yy")).toBe("Date d'échéance Mar'26");
    expect(applyFormat("2026-03-01", "date:l'an yyyy")).toBe("l'an 2026");
  });

  it("reads a pattern written in another script", () => {
    /**
     * The tokens are Latin and every other script's date words are not, so a
     * word rule of "any Unicode letter" — the first fix for the accent case
     * above — swept `年`, `월` and `г` into the run beside a token and made the
     * whole thing literal. `yyyy年MM月dd日` is the ordinary Japanese and Chinese
     * pattern; it printed itself, on every merged slide.
     *
     * Latin is the line because it is where the tokens live: a letter from
     * another script cannot be part of a word a token is hiding inside, so it
     * does not need to end one. The manual invites writing the month in your
     * own language, and this is what that means for languages not written in
     * this alphabet.
     */
    expect(applyFormat("2026-03-01", "date:yyyy年MM月dd日")).toBe("2026年03月01日");
    expect(applyFormat("2026-03-01", "date:yyyy년 MM월 dd일")).toBe("2026년 03월 01일");
    expect(applyFormat("2026-03-01", "date:dd MMMM yyyyг.")).toBe("01 March 2026г.");
    // The example the manual prints. A doc showing an output nobody checked is
    // the same defect as a stale comment.
    expect(applyFormat("2026-03-01", "date:yyyy年MM月dd日")).toBe("2026年03月01日");
  });

  it("leaves an ordinary word that begins with a token's letter alone", () => {
    /**
     * `den`, `dato`, `due` and `deadline` all start with a `d`, and `M` and `y`
     * begin plenty more. A rule that tokenizes the head of a run and prints its
     * tail as written turns `Berlin, den d. MMMM yyyy` — a perfectly ordinary
     * long-date pattern — into `Berlin, 1en 1. March 2026`, on every merged
     * slide.
     */
    expect(applyFormat("2026-03-01", "date:Berlin, den d. MMMM yyyy")).toBe("Berlin, den 1. March 2026");
    expect(applyFormat("2026-03-01", "date:dato d MMM yyyy")).toBe("dato 1 Mar 2026");
    expect(applyFormat("2026-03-01", "date:d MMM yyyy deadline")).toBe("1 Mar 2026 deadline");
    expect(applyFormat("2026-03-01", "date:MMMM d yyyy due")).toBe("March 1 2026 due");
  });

  it("leaves a weekday spelling alone rather than reading it as two days", () => {
    // `dddd` is how Excel and .NET spell a weekday name, and this add-in has no
    // such token. Read greedily it is `dd` twice and prints "0101" — a number
    // that means nothing, where the pattern printed as written at least shows
    // the author what was not understood.
    expect(applyFormat("2026-03-01", "date:dddd")).toBe("dddd");
    expect(applyFormat("2026-03-01", "date:dddd d MMM")).toBe("dddd 1 Mar");
    expect(applyFormat("2026-03-01", "date:yyyyyy")).toBe("yyyyyy");
    // The ODD-length repeats too, which the "same letter twice" rule alone does
    // not cover: `MMMMM` is `MMMM` plus a stray `M`, and printing `MarchM` is
    // the `MMM` → `MarM` failure this file's history already names.
    expect(applyFormat("2026-03-01", "date:MMMMM")).toBe("MMMMM");
    expect(applyFormat("2026-03-01", "date:MMMMMM")).toBe("MMMMMM");
    expect(applyFormat("2026-03-01", "date:ddd")).toBe("ddd");
    expect(applyFormat("2026-03-01", "date:yyy")).toBe("yyy");
    expect(applyFormat("2026-03-01", "date:yyyyy")).toBe("yyyyy");
  });
});
