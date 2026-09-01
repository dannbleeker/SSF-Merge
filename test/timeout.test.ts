import { describe, expect, it, vi } from "vitest";
import { BUDGET, Timeout, withTimeout } from "../src/host/timeout.js";
import { beginRun, traceLog } from "../src/core/trace.js";

describe("bounding a host call", () => {
  it("passes a result through when the work answers in time", async () => {
    await expect(withTimeout(Promise.resolve("ok"), 1000, "a quick read")).resolves.toBe("ok");
  });

  it("passes a rejection through unchanged", async () => {
    // A host error must not be dressed up as a timeout: they send a reader to
    // different files.
    await expect(withTimeout(Promise.reject(new Error("InvalidArgument")), 1000, "an insert")).rejects.toThrow(
      "InvalidArgument",
    );
  });

  it("gives up when the work does not answer, and names the call", async () => {
    const never = new Promise<string>(() => {});
    const err = await withTimeout(never, 5, "counting the deck's slides").catch((e: unknown) => e);
    expect(err).toBeInstanceOf(Timeout);
    expect((err as Timeout).what).toBe("counting the deck's slides");
    expect((err as Error).message).toContain("counting the deck's slides");
  });

  it("clears its timer when the work wins, so nothing is left pending", async () => {
    // A timer left running keeps a Node process alive after the work is done,
    // which turns a green suite into one that hangs. Counted rather than
    // reasoned about: the race must clear exactly the timer it set.
    const clear = vi.spyOn(globalThis, "clearTimeout");
    const before = clear.mock.calls.length;
    await withTimeout(Promise.resolve(1), 10000, "a quick read");
    expect(clear.mock.calls.length).toBeGreaterThan(before);
    clear.mockRestore();
  });

  it("names a rejection that is not an Error, rather than tracing [object Object]", async () => {
    /**
     * `err instanceof Error` is the only thing between a host rejection and the
     * run log. Office.js rejects with `OfficeExtension.Error`, which is one —
     * but a rejected string, a rejected plain object and a rejected `undefined`
     * are all things a promise can carry, and the run log is the only
     * diagnostic a task-pane user can hand over. A line reading
     * `raised call=an insert` with no error at all sends the reader nowhere.
     *
     * The three states this function reports — answered, gave up, raised — are
     * separate because they send a reader to different files. A raise with
     * nothing in it collapses the third back into "the host got in the way".
     */
    beginRun();
    await withTimeout(Promise.reject("GeneralException"), 1000, "an insert").catch(() => undefined);
    const raised = traceLog().entries.find((l) => l.message === "raised");
    expect(raised?.data?.error, "the run log is what a user hands over").toBe("GeneralException");

    beginRun();
    await withTimeout(Promise.reject({ code: 5010 }), 1000, "an insert").catch(() => undefined);
    const other = traceLog().entries.find((l) => l.message === "raised");
    expect(typeof other?.data?.error, "still a string, whatever was thrown").toBe("string");
  });

  it("clears its timer when the work FAILS too", async () => {
    const clear = vi.spyOn(globalThis, "clearTimeout");
    const before = clear.mock.calls.length;
    await withTimeout(Promise.reject(new Error("nope")), 10000, "an insert").catch(() => undefined);
    expect(clear.mock.calls.length).toBeGreaterThan(before);
    clear.mockRestore();
  });
});

describe("a call that answers after its run has ended", () => {
  /**
   * An abandoned call is not cancelled — nothing here can cancel a host call —
   * so it may still answer, and its `.then` still traces. If the pane has begun
   * another run by then, that line lands in the NEW run's log: the abandoned
   * run ends on an `issued` with no answer, reading as a call that never came
   * back, and the next run opens with an `answered` for a call nobody in it
   * made. Both are false, and a run log is the only diagnostic a task-pane user
   * can hand over.
   */
  it("does not write into the next run's log", async () => {
    // NESTED, which is the shape that reaches it. A single `withTimeout` whose
    // work answers late traces nothing — its race already settled on the
    // timeout and the value handler never runs. What survives the outer budget
    // is the INNER call's own wrapper, still waiting on its own larger one.
    beginRun();
    let settle: (v: string) => void = () => undefined;
    const held = new Promise<string>((resolve) => (settle = resolve));
    const inner = withTimeout(held, 10_000, "reading slide ids");
    // The outer one gives up first, exactly as a budget nested inside another
    // of its own size does.
    await expect(withTimeout(inner, 20, "reading the selection")).rejects.toThrow(Timeout);

    // The pane moves on, and only then does the inner call come back.
    beginRun();
    settle("late");
    await inner;
    await Promise.resolve();

    const lines = traceLog().entries.map((e) => `${e.message} ${typeof e.data?.call === "string" ? e.data.call : ""}`);
    expect(lines, "a line from a run that has ended").not.toContain("answered reading slide ids");
    expect(lines, "and nothing else leaked either").toEqual([]);
  });

  it("still records an answer inside its own run", async () => {
    // The other direction: the stamp must not swallow an ordinary line.
    beginRun();
    await withTimeout(Promise.resolve("ok"), 1000, "a quick read");
    expect(traceLog().entries.map((e) => e.message)).toEqual(["issued", "answered"]);
  });
});

describe("the budgets", () => {
  it("gives an insert more room than a read", () => {
    // Thirty seconds was measured too short once: the probe's third sheet timed
    // out on an insert whose deck delta showed both slides had landed.
    expect(BUDGET.insert).toBeGreaterThan(BUDGET.read);
    expect(BUDGET.insert).toBeGreaterThanOrEqual(60000);
  });

  it("gives reading the whole package the most room of all", () => {
    // It arrives slice by slice, and a large deck is a lot of slices.
    expect(BUDGET.file).toBeGreaterThanOrEqual(BUDGET.insert);
  });
});
