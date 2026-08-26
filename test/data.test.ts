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
    const rs = toRecordSet([["Name", ""], ["Ada", "x"]]);
    expect(rs.columns.map((c) => c.name)).toEqual(["Name", "Column 2"]);
    expect(rs.rows[0]?.["Column 2"]).toBe("x");
  });

  it("keeps both of two columns with the same header", () => {
    const rs = toRecordSet([["Name", "Name"], ["a", "b"]]);
    expect(rs.columns.map((c) => c.name)).toEqual(["Name", "Name 2"]);
  });

  it("drops a wholly blank row and keeps a partly filled one", () => {
    const rs = toRecordSet([["A", "B"], ["", ""], ["x", ""]]);
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
