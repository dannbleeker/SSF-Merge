import { describe, expect, it } from "vitest";
import { applyFormat, formatNumber, numericValue, parseDate } from "../src/core/data/format.js";
import { detectType, looksLikeDate, parseDelimited, toRecordSet } from "../src/core/data/recordset.js";

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
