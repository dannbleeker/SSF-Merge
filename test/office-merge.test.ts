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
const { Pkg } = await import("../src/core/pptx/pkg.js");
const { makeDeck } = await import("./fixtures/deck.js");

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

describe("the whole-deck route sends only the merged slides", () => {
  /**
   * The defect this describes was live, and it is the one a real-host round on
   * an older PowerPoint would have found the expensive way.
   *
   * `readTemplate` has two routes. On `subset` — PowerPointApi 1.10 and up —
   * `exportAsBase64Presentation` hands back a package holding ONLY the template
   * block, and `templateOffset` is 0. On `file` — everything below 1.10, which
   * this add-in supports, since its floor is 1.2 — `getFileAsync` hands back
   * the USER'S ENTIRE PRESENTATION, and the offset says where in it the block
   * begins.
   *
   * `runMerge` removed `prepared.block.slides` and nothing else, so on the file
   * route the package handed to `insertSlidesFromBase64` was the user's whole
   * deck, minus the template block, plus the clones. Merging three rows into a
   * forty-slide deck would have inserted forty-six slides: a second copy of
   * everything the user had.
   *
   * It is not silent — `insertVerdict` compares the deck delta against
   * `expected` and would report the mismatch — but by then the slides are in
   * the deck, and "the merge duplicated my presentation" is not a diagnosis
   * anyone should have to make from a verdict line.
   */
  const rows = {
    columns: [{ name: "First", type: "text" as const }],
    rows: [{ First: "Ada" }, { First: "Grace" }],
  };

  /** A five-slide deck whose block is slide 3 — the file route's shape. */
  async function wholeDeck(): Promise<string> {
    const bytes = await makeDeck([
      { paragraphs: [["Title"]] },
      { paragraphs: [["Agenda"]] },
      { paragraphs: [["Hello ", "{{First}}"]] },
      { paragraphs: [["Appendix"]] },
      { paragraphs: [["Thanks"]] },
    ]);
    return Buffer.from(bytes).toString("base64");
  }

  it("sends two slides, not the user's whole deck plus two", async () => {
    // offset 2 = the block starts at the third slide of the package, which is
    // exactly what `templateOffset("file", 2)` answers.
    host.readTemplate.mockResolvedValueOnce({ base64: await wholeDeck(), offset: 2 });
    host.insertDeck.mockResolvedValueOnce({ verdict: "yes", detail: "landed", landed: 2, before: 5, after: 7 });

    const out = await runMerge({ from: 3, to: 3, records: rows });
    expect(out.ok).toBe(true);

    const sent = host.insertDeck.mock.calls[0]?.[0] as string;
    const pkg = await Pkg.open(sent);
    // The package's OWN count, not what the caller believed it built. The four
    // slides the user already had must not be in it.
    expect(await pkg.slidePaths()).toHaveLength(2);
  });

  it("tells the host how many slides the package HOLDS", async () => {
    // `expected` is what `insertVerdict` grades the deck delta against, so a
    // number taken from anywhere but the package is a verdict about the wrong
    // thing. Measure the artefact you hand over, not your intent.
    host.readTemplate.mockResolvedValueOnce({ base64: await wholeDeck(), offset: 2 });
    host.insertDeck.mockResolvedValueOnce({ verdict: "yes", detail: "landed", landed: 2, before: 5, after: 7 });

    await runMerge({ from: 3, to: 3, records: rows });

    const sent = host.insertDeck.mock.calls[0]?.[0] as string;
    const expected = host.insertDeck.mock.calls[0]?.[1] as number;
    expect(expected).toBe((await (await Pkg.open(sent)).slidePaths()).length);
  });

  it("still sends only the clones on the subset route", async () => {
    // The route that already worked, so the fix cannot have been a wash: here
    // the package IS the block, and removing it leaves the clones alone.
    const bytes = await makeDeck([{ paragraphs: [["Hello ", "{{First}}"]] }]);
    host.readTemplate.mockResolvedValueOnce({ base64: Buffer.from(bytes).toString("base64"), offset: 0 });
    host.insertDeck.mockResolvedValueOnce({ verdict: "yes", detail: "landed", landed: 2, before: 3, after: 5 });

    await runMerge({ from: 1, to: 1, records: rows });

    const sent = host.insertDeck.mock.calls[0]?.[0] as string;
    expect(await (await Pkg.open(sent)).slidePaths()).toHaveLength(2);
  });
});

describe("a torn insert is reported in rows", () => {
  /** Three rows, two template slides each — six slides expected. */
  const rows = {
    columns: [{ name: "First", type: "text" as const }],
    rows: [{ First: "Ada" }, { First: "Grace" }, { First: "Katherine" }],
  };

  async function block(): Promise<string> {
    const bytes = await makeDeck([{ paragraphs: [["Hello ", "{{First}}"]] }, { paragraphs: [["Bye ", "{{First}}"]] }]);
    return Buffer.from(bytes).toString("base64");
  }

  it("names the ROW that is incomplete, not the slide count", () => {
    // "5 of 6 slides landed" is true and useless: rows 1 and 2 look perfect
    // and the user finds the short one at the end of the deck.
    return (async () => {
      host.readTemplate.mockResolvedValueOnce({ base64: await block(), offset: 0 });
      host.insertDeck.mockResolvedValueOnce({
        verdict: "no",
        detail: "5 of 6 slide(s) landed",
        landed: 5,
        before: 2,
        after: 7,
      });

      const out = await runMerge({ from: 1, to: 2, records: rows });
      expect(out.ok).toBe(false);
      expect(out.detail).toContain("2 of 3 row(s) landed complete");
      expect(out.detail).toContain("row 3 got 1 of its 2 slide(s)");
      // The action, which does not depend on WHICH row tore.
      expect(out.detail).toMatch(/take the slides back/i);
    })();
  });

  it("carries the row counts for anything else that reports", async () => {
    host.readTemplate.mockResolvedValueOnce({ base64: await block(), offset: 0 });
    host.insertDeck.mockResolvedValueOnce({ verdict: "no", detail: "5 of 6", landed: 5, before: 2, after: 7 });

    const out = await runMerge({ from: 1, to: 2, records: rows });
    expect(out).toMatchObject({ rowsComplete: 2, rowsTorn: 1, rowsAbsent: 0 });
  });

  it("still says the host took NOTHING when it took nothing", async () => {
    // A refusal and a tear are different failures. Reading a refusal as
    // "0 of 3 rows landed complete" buries the fact that the call was rejected.
    host.readTemplate.mockResolvedValueOnce({ base64: await block(), offset: 0 });
    host.insertDeck.mockResolvedValueOnce({
      verdict: "no",
      detail: "the call raised nothing and the deck did not grow",
      landed: 0,
      before: 2,
      after: 2,
    });

    const out = await runMerge({ from: 1, to: 2, records: rows });
    expect(out.detail).toContain("did not take it");
    expect(out.detail).not.toMatch(/row\(s\) landed complete/);
  });

  it("says nothing about rows when every row landed", async () => {
    // 240 of 240 rows is noise beside "720 slides added".
    host.readTemplate.mockResolvedValueOnce({ base64: await block(), offset: 0 });
    host.insertDeck.mockResolvedValueOnce({ verdict: "yes", detail: "all 6 landed", landed: 6, before: 2, after: 8 });

    const out = await runMerge({ from: 1, to: 2, records: rows });
    expect(out.ok).toBe(true);
    expect(out).toMatchObject({ rowsComplete: 3, rowsTorn: 0, rowsAbsent: 0 });
    expect(out.detail).not.toMatch(/row\(s\) landed complete/);
  });
});
