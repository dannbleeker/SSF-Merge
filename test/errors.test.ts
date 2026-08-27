import { describe, expect, it } from "vitest";

import { ERROR_CHARS, readable, short } from "../src/host/errors.js";

describe("error text that reaches a user is bounded", () => {
  it("caps a message and counts what it dropped", () => {
    // Not a bare ellipsis: a reader has to be able to tell a truncated
    // sentence from one that ended oddly, and to know whether what is missing
    // was a clause or a megabyte.
    const huge = `PowerPoint said no: ${"A".repeat(200_000)}`;
    const out = readable(new Error(huge));
    expect(out.length).toBeLessThan(ERROR_CHARS + 60);
    expect(out).toContain("PowerPoint said no:");
    expect(out).toMatch(/\(\d+ more characters\)/);
  });

  it("leaves an ordinary sentence exactly as it is", () => {
    // Every message that reaches here is already a sentence somebody wrote.
    expect(readable(new Error("PowerPoint would not name every slide."))).toBe(
      "PowerPoint would not name every slide.",
    );
  });

  it("handles a throw that is not an Error", () => {
    expect(readable("just a string")).toBe("just a string");
    expect(readable(undefined)).toBe("undefined");
  });

  it("does not cut a message that is exactly at the limit", () => {
    const exact = "A".repeat(ERROR_CHARS);
    expect(short(exact)).toBe(exact);
  });
});
