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

/**
 * A host error object whose `message` is not a string.
 *
 * Office.js failures are routinely plain objects rather than `Error`s, and the
 * message in one is the whole point — so `readable` reaches for `e.message`,
 * guarded by `typeof e.message === "string"`. Loosening that guard to accept
 * any object that merely HAS a `message` left the suite green, which
 * `scripts/mutate-core.mjs` found.
 *
 * It is the one function in this file that must never throw. It runs when
 * something has already gone wrong, and what it returns is the only sentence
 * the user gets: a `TypeError` from `short(5010)` replaces the pane's report of
 * the failure with a second failure nobody can read.
 */
describe("a host error whose message is not a string", () => {
  const odd: [string, unknown][] = [
    ["a number, which is what a bare error code looks like", { name: "GeneralException", message: 5010, code: 5010 }],
    ["a nested object", { message: { text: "inner" }, code: 1 }],
    ["null", { message: null, name: "X" }],
    ["an array", { message: ["a", "b"] }],
  ];

  it.each(odd)("describes it instead of throwing: %s", (_label, e) => {
    let out: string;
    expect(() => {
      out = readable(e);
    }, "the sentence the user gets must not itself fail").not.toThrow();
    out = readable(e);
    expect(out).not.toContain("[object Object]");
    expect(out.length).toBeGreaterThan(0);
    // The shape is named, which is a fact, where a stringification is not.
    expect(out).toContain("message");
  });

  it("still prefers a real message when there is one", () => {
    // The other direction: tightening the guard must not lose the case it
    // exists for, which is every ordinary Office.js failure.
    expect(readable({ name: "GeneralException", message: "InvalidParam passed to GetItem(id)" })).toBe(
      "InvalidParam passed to GetItem(id)",
    );
  });
});

describe("a raise with nothing in it", () => {
  /**
   * An `OfficeExtension.Error` routinely carries an empty `message` and puts
   * the content in `debugInfo`. `readable` answered "" for it, so the sentence
   * the pane builds around the answer stopped mid-air: "The merge did not
   * run: ". An empty answer is the same defect as "[object Object]", which is
   * what the rest of this function exists to refuse — it occupies the space
   * where a reason goes and says nothing.
   */
  it("says so rather than answering with an empty string", () => {
    expect(readable(new Error(""))).toBe("the host raised nothing this pane can describe.");
    expect(readable("")).toBe("the host raised nothing this pane can describe.");
  });

  it("falls through to the shape when an object's message is empty", () => {
    // The object may still carry a name and a code, and "InvalidArgument /
    // 5010" is something a reader can repeat to somebody who can act on it.
    const said = readable({ name: "InvalidArgument", message: "", code: 5010 });
    expect(said).toContain("InvalidArgument");
    expect(said).toContain("5010");
  });

  it("still prefers a message that has something in it", () => {
    // The other direction: the guard must not cost the ordinary case.
    expect(readable(new Error("the deck is open elsewhere"))).toBe("the deck is open elsewhere");
  });

  it("does not raise when the error's own message raises", () => {
    // `readable` is the `whenItRaises` handler, so it runs inside a catch — a
    // throw of its own escapes `duringRun` and leaves the pane holding
    // "Merging…" for ever, which is the state the caller exists to survive.
    // Reading `.message` is the risk: it is a property, and a getter may throw.
    const hostile = {
      get message(): string {
        throw new Error("boom");
      },
    };
    expect(() => readable(hostile)).not.toThrow();
    expect(readable(hostile)).toContain("could not read");

    // And the same shape as a real Error, which takes the earlier branch.
    const asError = Object.create(Error.prototype) as Error;
    Object.defineProperty(asError, "message", {
      get(): string {
        throw new Error("boom");
      },
    });
    expect(() => readable(asError)).not.toThrow();
  });
});
