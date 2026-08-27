/**
 * The merge seam's REFUSALS, which are the half that had no test.
 *
 * `src/office` cannot run in the suite — it calls Office.js — but the two
 * functions the pane awaits are ordinary async functions over an injectable
 * boundary: everything below `readTemplate` is the engine. Mocking that one
 * call is enough to drive every path the pane can be handed, and those paths
 * are where an adversarial review found a rejection escaping into the pane with
 * nothing to catch it.
 *
 * The rule this file exists for: a refusal here is an OUTCOME, never a raise.
 * The pane awaits both of these from a click handler, so a rejection is an
 * unhandled one, and the user is left looking at a button that never comes back.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";

const host = vi.hoisted(() => ({
  readTemplate: vi.fn<(b: { from: number; to: number }) => Promise<{ base64: string; offset: number }>>(),
  slideCount: vi.fn<() => Promise<number>>(),
  insertDeck: vi.fn(),
  undoInsert: vi.fn(),
}));

vi.mock("../src/office/powerpoint.js", () => host);

const { inspectBlock, runMerge } = await import("../src/office/merge.js");

const records = { columns: [{ name: "First", type: "text" as const }], rows: [{ First: "Ada" }] };

beforeEach(() => {
  host.readTemplate.mockReset();
  host.slideCount.mockReset().mockResolvedValue(3);
  host.insertDeck.mockReset();
});

describe("inspectBlock answers rather than raising", () => {
  it("turns readTemplate's refusal into an outcome", () => {
    // `readTemplate` throws by design: `blockIds` produced a sentence and a
    // throw is how it leaves a PowerPoint.run callback.
    host.readTemplate.mockRejectedValueOnce(new Error("PowerPoint would not name every slide."));
    return expect(inspectBlock({ from: 4, to: 6 })).resolves.toMatchObject({
      ok: false,
      detail: "PowerPoint would not name every slide.",
    });
  });

  it("turns a package it cannot OPEN into an outcome", async () => {
    // The try covered `readTemplate` alone at first, and `Pkg.open` was awaited
    // past it. A host that answers the export with bytes JSZip cannot read then
    // rejected straight through `useBlock()` into `void`, leaving the pane
    // saying "Reading the slides…" for the rest of the session.
    host.readTemplate.mockResolvedValueOnce({ base64: "bm90IGEgemlwIGF0IGFsbA==", offset: 0 });
    const report = await inspectBlock({ from: 1, to: 1 });
    expect(report.ok).toBe(false);
    expect(report.detail.length).toBeGreaterThan(0);
  });

  it("turns a package with no presentation part into an outcome", async () => {
    // A zip that opens and is not a deck: `prepareBlock` reaches
    // `pkg.slidePaths()`, which throws by name for the missing part. Also past
    // the original try.
    const JSZip = (await import("jszip")).default;
    const zip = new JSZip();
    zip.file("hello.txt", "not a presentation");
    const base64 = await zip.generateAsync({ type: "base64" });
    host.readTemplate.mockResolvedValueOnce({ base64, offset: 0 });
    const report = await inspectBlock({ from: 1, to: 1 });
    expect(report.ok).toBe(false);
    expect(report.detail).toContain("presentation.xml");
  });
});

describe("runMerge answers rather than raising", () => {
  it("turns readTemplate's refusal into an outcome, keeping deckAtStart", async () => {
    // The pane awaits this from a click handler. A rejection here was
    // unhandled, and it took `deckAtStart` with it — the number a positional
    // undo is clamped against, which merge()'s own docstring says must be held
    // before anything is shown.
    host.readTemplate.mockRejectedValueOnce(new Error("PowerPoint would not name every slide."));
    const outcome = await runMerge({ from: 4, to: 6, records });
    expect(outcome.ok).toBe(false);
    expect(outcome.detail).toContain("would not name");
    expect(outcome.deckAtStart, "the undo clamp survives the refusal").toBe(3);
    expect(outcome.added).toBe(0);
  });

  it("turns an unreadable package into an outcome too", async () => {
    host.readTemplate.mockResolvedValueOnce({ base64: "bm90IGEgemlw", offset: 0 });
    const outcome = await runMerge({ from: 1, to: 1, records });
    expect(outcome.ok).toBe(false);
    expect(outcome.deckAtStart).toBe(3);
  });

  it("still refuses an empty row set before spending a host read", async () => {
    const outcome = await runMerge({ from: 1, to: 1, records: { columns: [], rows: [] } });
    expect(outcome.detail).toContain("no rows");
    expect(host.readTemplate).not.toHaveBeenCalled();
  });
});
