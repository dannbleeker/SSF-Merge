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
  });

  it("keeps the message out of an object that is not an Error", () => {
    /**
     * `String(e)` reaches Object's default stringification, so a thrown
     * `{ message: "InvalidArgument", code: 5 }` arrived as "[object Object]" —
     * discarding the message inside it, in the sentence the user is given to
     * explain a failed merge. An Office.js async failure is routinely a plain
     * object carrying exactly those fields.
     *
     * `formatValue` in `trace.ts` already refuses to print "[object Object]",
     * with the reasoning beside it: "a line that occupies space and answers
     * nothing". This is the same rule on the path that reaches a person.
     */
    expect(readable({ message: "InvalidArgument", code: 5 })).toBe("InvalidArgument");
    // No message: the SHAPE, which a reader can at least repeat to somebody.
    expect(readable({ code: 5 })).toBe('{"code":5}');
  });

  it("says a raise had nothing in it, rather than the word undefined", () => {
    // This replaces an assertion that pinned `readable(undefined)` to the
    // string "undefined". That is a JavaScript value name reaching a user as
    // the whole explanation of why their merge failed.
    expect(readable(undefined)).toBe("the host raised nothing this pane can describe.");
    expect(readable(null)).toBe("the host raised nothing this pane can describe.");
  });

  it("names a circular object rather than raising while reporting a raise", () => {
    // An Office error object can carry a cycle, and `JSON.stringify` throws on
    // one. A reporter that raises while reporting a raise loses both the
    // original failure and the pane, so this branch answers instead.
    const cyclic: Record<string, unknown> = { code: 5 };
    cyclic.self = cyclic;
    expect(readable(cyclic)).toBe("the host raised something this pane could not read.");
  });

  it("names a thrown function instead of printing its source at the user", () => {
    // `String(fn)` prints the whole body into the sentence. This is a caller's
    // slip rather than a host failure, and it says so — `formatValue` in
    // `trace.ts` gives the same answer for the same reason.
    expect(readable(() => 1)).toBe("the host raised a function, which is a bug in this add-in.");
  });

  it("reads the primitives a throw can be, each as itself", () => {
    /**
     * One `String()` over "everything left" is the call that put
     * "[object Object]" on screen in the first place, so each of these is named
     * by its own branch. They were reachable and untested: this file is on the
     * path of every failure the pane reports, and was its worst-covered by a
     * wide margin.
     */
    expect(readable(0)).toBe("0");
    expect(readable(false)).toBe("false");
    expect(readable(10n)).toBe("10");
    expect(readable(Symbol("InvalidArgument"))).toBe("Symbol(InvalidArgument)");
  });

  it("caps every shape, not only an Error's message", () => {
    // `short` is applied per branch rather than once at the exit, so a branch
    // added without it is a branch with no cap. A 200k string thrown bare is
    // the case that reaches the pane through `state.notice` to a DOM node.
    const huge = "B".repeat(200_000);
    expect(readable(huge).length).toBeLessThan(ERROR_CHARS + 60);
    expect(readable({ message: huge }).length).toBeLessThan(ERROR_CHARS + 60);
    expect(readable({ note: huge }).length).toBeLessThan(ERROR_CHARS + 60);
  });

  it("does not cut a message that is exactly at the limit", () => {
    const exact = "A".repeat(ERROR_CHARS);
    expect(short(exact)).toBe(exact);
  });
});
