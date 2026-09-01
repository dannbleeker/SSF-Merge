import { afterEach, describe, expect, it, vi } from "vitest";
import { ERROR_CHARS } from "../src/host/errors.js";
import { installFakeHost, uninstallFakeHost } from "./fixtures/host.js";

/**
 * The one call the pane makes before it has decided anything.
 *
 * `src/office` mostly cannot run in the suite — it calls Office.js — but
 * `hostSupports` is a single expression over `Office.context`, and what it does
 * when that object is not what it expects decides whether an unsupported host
 * gets a sentence or a blank pane.
 *
 * `ready()` uses it to render "this PowerPoint is too old", and
 * `hostEnvironment()` uses it to say which build was refused. Both run before
 * anything else, so a raise from either leaves the user with nothing to read on
 * exactly the host that needed the message — and nothing to report, because the
 * pane never drew.
 */
async function supportsWith(context: unknown): Promise<(v: string) => boolean> {
  vi.resetModules();
  (globalThis as unknown as { Office: unknown }).Office = { context };
  const mod = await import("../src/office/powerpoint.js");
  return mod.hostSupports;
}

afterEach(() => {
  delete (globalThis as unknown as { Office?: unknown }).Office;
});

describe("asking the host what it supports", () => {
  it("answers what the host says when the host can answer", () => {
    // The ordinary path, so the guard below is not just "always false".
    const asked: string[] = [];
    return supportsWith({
      requirements: {
        isSetSupported: (name: string, v: string) => {
          asked.push(`${name} ${v}`);
          return v === "1.2";
        },
      },
    }).then((supports) => {
      expect(supports("1.2")).toBe(true);
      expect(supports("1.10")).toBe(false);
      expect(asked).toEqual(["PowerPointApi 1.2", "PowerPointApi 1.10"]);
    });
  });

  it("answers false rather than throwing when there is no requirements object", async () => {
    const supports = await supportsWith({});
    expect(supports("1.2")).toBe(false);
  });

  it("answers false rather than throwing when there is no context at all", async () => {
    const supports = await supportsWith(undefined);
    expect(supports("1.2")).toBe(false);
  });

  it("answers false rather than throwing when the host raises", async () => {
    // A host that has the method and refuses the question. Same treatment: it
    // cannot answer, so it does not support it.
    const supports = await supportsWith({
      requirements: {
        isSetSupported: () => {
          throw new Error("no");
        },
      },
    });
    expect(supports("1.2")).toBe(false);
  });
});

/**
 * The two calls that ACT on the user's deck, driven against a fake host.
 *
 * `test/office-merge.test.ts` mocks both of these wholesale, so what they do
 * with a host's answer was checked by nothing. Both defects below were found
 * that way and neither is visible to a mock of the function holding it.
 */
describe("what an insert and an undo say when the host misbehaves", () => {
  async function host(options: Parameters<typeof installFakeHost>[0] = {}) {
    vi.resetModules();
    const fake = installFakeHost(options);
    const module = await import("../src/office/powerpoint.js");
    return { fake, module };
  }

  afterEach(() => {
    uninstallFakeHost();
  });

  it("bounds what the host said, so a failed insert cannot put the deck on screen", async () => {
    // Office echoes an argument back through `debugInfo`, and the argument to
    // `insertSlidesFromBase64` is the ENTIRE merged deck as base64. Uncapped,
    // this sentence carried megabytes of the user's own merged rows to the
    // pane — and into a bug report, if they pasted it. `errors.ts` exists for
    // this and its callers were the paths that REJECT; this one is caught and
    // returned, so it never asked.
    const { module } = await host({
      onInsert: () => {
        throw new Error(`InvalidArgument: ${"A".repeat(2_000_000)}`);
      },
    });
    const outcome = await module.insertDeck("PKG", 2);
    expect(outcome.detail.length).toBeLessThanOrEqual(ERROR_CHARS + 64);
    expect(outcome.detail, "the reason itself is gone").toContain("InvalidArgument");
    expect(outcome.detail, "and what was dropped is counted rather than elided").toMatch(/more characters/);
  });

  it("names the slides the way the rail numbers them, never a 0-based index", async () => {
    // The pane speaks the thumbnail rail's numbering and nothing else. This
    // sentence reaches the user verbatim — "Nothing was removed — ${detail}" —
    // and said "from index 3" for slides the rail calls 4 to 9: a number one
    // before the ones actually touched, in a sentence about slides being
    // deleted. The refusal branch a few lines above it already got this right.
    const { fake, module } = await host({
      slides: ["u1", "u2", "u3", "m1", "m2", "m3", "m4", "m5", "m6"],
    });
    const outcome = await module.undoInsert(3, 6, "run-1");
    expect(outcome.removed).toBe(6);
    expect(outcome.detail, "an internal index reached the user").not.toContain("index");
    expect(outcome.detail).toContain("slides 4 to 9");
    expect(fake.slides).toEqual(["u1", "u2", "u3"]);
  });
});
