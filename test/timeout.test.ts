import { describe, expect, it, vi } from "vitest";
import { BUDGET, Timeout, withTimeout } from "../src/host/timeout.js";

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

  it("clears its timer when the work FAILS too", async () => {
    const clear = vi.spyOn(globalThis, "clearTimeout");
    const before = clear.mock.calls.length;
    await withTimeout(Promise.reject(new Error("nope")), 10000, "an insert").catch(() => undefined);
    expect(clear.mock.calls.length).toBeGreaterThan(before);
    clear.mockRestore();
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
