import { afterEach, describe, expect, it, vi } from "vitest";

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
