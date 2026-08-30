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
    ];
    const meaning = new Map<string, number>();
    for (const line of languages) {
      line.split(" ").forEach((word, i) => {
        const month = i + 1;
        const already = meaning.get(word);
        expect(already ?? month, `${word} means two different months`).toBe(month);
        meaning.set(word, month);
        // And the engine agrees with the list.
        expect(parseDate(`1 ${word} 2026`)?.getUTCMonth(), `1 ${word} 2026`).toBe(i);
      });
    }
    expect(meaning.size).toBe(25);
  });

  it("still refuses a day that month does not have, whatever the language", () => {
    // The table must not become a way round the guard added with it.
    expect(parseDate("31 apr 2026")).toBeUndefined();
    expect(parseDate("30 februar 2026")).toBeUndefined();
    expect(parseDate("32 maj 2026")).toBeUndefined();
  });

  it("keeps whatever the platform already handled", () => {
    /**
     * The fallback stays on purpose. It is what makes French and Italian month
     * names work today for users this table does not list, and dropping it
     * would turn a partial answer into no answer for them. Its inconsistency is
     * why the table exists — it is a floor, not the mechanism.
     */
    expect(parseDate("1 janvier 2026")?.toISOString().slice(0, 10)).toBe("2026-01-01");
    expect(parseDate("1 March 2026")?.toISOString().slice(0, 10)).toBe("2026-03-01");
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
    const seps = ["", ",", ".", " "];
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
    expect(CORPUS.length, "the sweep stopped covering anything").toBeGreaterThan(5000);
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
    // And the `new Date` fallback, which is the branch the guard now feeds.
    expect(applyFormat("1 Mar 2026", "date:d MMM yyyy")).toBe("1 Mar 2026");
  });
});
