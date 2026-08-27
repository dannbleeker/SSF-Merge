import { describe, expect, it, beforeEach, vi } from "vitest";

import {
  MAX_ENTRIES,
  beginRun,
  elapsed,
  formatTraceLine,
  formatValue,
  onTrace,
  trace,
  traceLog,
  traceText,
  tracing,
} from "../src/core/trace.js";
import { BUDGET, Timeout, withTimeout } from "../src/host/timeout.js";

beforeEach(() => {
  onTrace(undefined);
  beginRun();
});

describe("the run record", () => {
  it("records nothing before a run begins", () => {
    // Not a performance concern — it is a correctness one. A module-level
    // buffer that fills before `beginRun` pairs one run's lines with another
    // run's numbers, which is the wrong turn that costs an hour.
    onTrace(undefined);
    // Simulate a fresh module: end the run by never having started one.
    const before = traceLog().entries.length;
    expect(before).toBe(0);
    expect(tracing()).toBe(true);
  });

  it("clears whatever came before when a run starts", () => {
    trace("merge", "first run");
    beginRun();
    expect(traceLog().entries).toEqual([]);
  });

  it("copies the payload, so a caller cannot rewrite history", () => {
    // A log is a record of what was true AT THE MOMENT it was written. Holding
    // the caller's object lets a later mutation change a file somebody may
    // already be reading as fact.
    const payload = { rows: 3 };
    trace("merge", "planned", payload);
    payload.rows = 999;
    expect(traceLog().entries[0]?.data).toEqual({ rows: 3 });
  });

  it("hands readers copies too", () => {
    trace("merge", "planned", { rows: 3 });
    const first = traceLog();
    (first.entries[0]?.data as { rows: number }).rows = 999;
    expect(traceLog().entries[0]?.data).toEqual({ rows: 3 });
  });

  it("drops the oldest and counts what it dropped", () => {
    for (let i = 0; i < MAX_ENTRIES + 5; i++) trace("merge", `step ${i}`);
    const log = traceLog();
    expect(log.entries).toHaveLength(MAX_ENTRIES);
    expect(log.dropped).toBe(5);
    // The oldest went, not the newest: what anyone reads in a run that stalled
    // is the end of it.
    expect(log.entries[0]?.message).toBe("step 5");
    expect(traceText()).toContain("(5 earlier line(s) dropped)");
  });

  it("never lets a broken window cost the record", () => {
    // The watcher is a renderer. It is at its most likely to throw during the
    // failure it exists to photograph.
    onTrace(() => {
      throw new Error("the pane is gone");
    });
    expect(() => {
      trace("merge", "still recorded");
    }).not.toThrow();
    expect(traceLog().entries).toHaveLength(1);
  });

  it("gives every line in one run the same clock origin", () => {
    const at = elapsed();
    expect(at).not.toBeNull();
    expect(at).toBeGreaterThanOrEqual(0);
    trace("merge", "a");
    expect(traceLog().entries[0]?.ms).toBeGreaterThanOrEqual(0);
  });
});

describe("a payload value on a line", () => {
  it("keeps 0, false and null, which are answers", () => {
    // `filled=0` is the ENTIRE finding of a merge that matched no placeholders.
    // Dropping falsy values would drop the one line worth reading.
    expect(formatValue(0)).toBe("0");
    expect(formatValue(false)).toBe("false");
    expect(formatValue(null)).toBe("null");
  });

  it("leaves off the two that mean nothing to say", () => {
    expect(formatValue(undefined)).toBeUndefined();
    expect(formatValue("")).toBeUndefined();
  });

  it("caps a value rather than banning a type", () => {
    // A first version elsewhere dropped everything whose typeof was "object",
    // which is every array — and the arrays were the numbers worth having.
    expect(formatValue([1, 2, 3])).toBe("[1,2,3]");
    const long = formatValue("x".repeat(500));
    expect(long?.length).toBeLessThan(500);
    expect(long?.endsWith("…")).toBe(true);
  });

  it("renders a shape rather than [object Object]", () => {
    const circular: Record<string, unknown> = {};
    circular.self = circular;
    expect(formatValue(circular)).toBe("{…}");
  });

  it("formats a line the same way for the screen and the file", () => {
    // One formatter, or the two will describe the same run differently.
    const line = formatTraceLine({ ms: 1234, scope: "host", message: "issued", data: { call: "x", budget: 15000 } });
    expect(line).toContain("host");
    expect(line).toContain("issued");
    expect(line).toContain("call=x");
    expect(line).toContain("budget=15000");
  });
});

describe("every host call names itself, on both populations", () => {
  const calls = () => traceLog().entries.filter((e) => e.scope === "host");

  it("records a call that WORKED, not only one that failed", () => {
    // The defect this repo shipped: `what` reached exactly one place, the
    // message of a Timeout nobody sees unless the call failed. A value
    // recorded only on failures cannot be compared against anything and is not
    // yet a measurement — so there was no baseline against which a 41-second
    // insert is normal or alarming.
    return withTimeout(Promise.resolve("ok"), BUDGET.read, "counting the deck's slides").then(() => {
      expect(calls().map((e) => e.message)).toEqual(["issued", "answered"]);
      expect(calls()[0]?.data).toMatchObject({ call: "counting the deck's slides", budget: BUDGET.read });
      expect(calls()[1]?.data?.ms).toBeTypeOf("number");
    });
  });

  it("writes the issued line BEFORE the call, so a call that never answers is on the record", async () => {
    // The ordering is the whole point, and it is why the line is named for
    // what it KNOWS. A sibling project's per-batch line was called "batch
    // committed" and written one statement before its sync, so every stall
    // left a line claiming the batch that killed it had committed.
    let release: (v: string) => void = () => {};
    const pending = new Promise<string>((res) => {
      release = res;
    });
    const race = withTimeout(pending, 10_000, "inserting the merged deck");
    expect(calls().map((e) => e.message)).toEqual(["issued"]);
    release("done");
    await race;
    expect(calls().map((e) => e.message)).toEqual(["issued", "answered"]);
  });

  it("tells a budget running out apart from a raise", async () => {
    vi.useFakeTimers();
    try {
      const never = new Promise<never>(() => {});
      const race = withTimeout(never, 50, "inserting the merged deck");
      const settled = expect(race).rejects.toBeInstanceOf(Timeout);
      await vi.advanceTimersByTimeAsync(60);
      await settled;
    } finally {
      vi.useRealTimers();
    }
    expect(calls().map((e) => e.message)).toEqual(["issued", "gave up waiting"]);
    // No `error` on a timeout: there was no error, the host simply never
    // answered, and inventing one would be a fact about us.
    expect(calls()[1]?.data).not.toHaveProperty("error");
  });

  it("carries the raise's own sentence when the call raised", async () => {
    await expect(withTimeout(Promise.reject(new Error("InvalidArgument")), 5000, "exporting")).rejects.toThrow();
    expect(calls().map((e) => e.message)).toEqual(["issued", "raised"]);
    expect(calls()[1]?.data?.error).toBe("InvalidArgument");
  });
});
