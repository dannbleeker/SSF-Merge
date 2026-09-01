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

describe("what became of the pictures", () => {
  /**
   * `runPlan` has always answered this and this seam threw it away, so every
   * field on `ImageOutcome` carried a comment about being named rather than
   * left silent and the only readers were tests. The pane's pre-merge tally
   * cannot stand in for it: that matches file NAMES and never opens one, so a
   * folder of renamed files passes it and places nothing.
   */
  it("carries it out of the merge, so something can say so", async () => {
    const bytes = await makeDeck([{ paragraphs: [["{{Photo|image}}"]] }]);
    host.readTemplate.mockResolvedValueOnce({ base64: Buffer.from(bytes).toString("base64"), offset: 0 });
    host.insertDeck.mockResolvedValueOnce({ verdict: "yes", detail: "landed", landed: 1, before: 3, after: 4 });

    const outcome = await runMerge({
      from: 1,
      to: 1,
      records: { columns: [{ name: "Photo", type: "image" as const }], rows: [{ Photo: "ada.png" }] },
      // Named like a picture and not one. Nothing before the merge reads a
      // byte, so this is the case only the outcome can report.
      images: new Map([["ada.png", new Uint8Array([1, 2, 3, 4])]]),
    });

    expect(outcome.pictures, "the merge said nothing about the pictures").toBeDefined();
    expect(outcome.pictures).toMatchObject({ placed: 0, unreadable: ["Photo"] });
    // The chart tally rides out on the same object and had the same problem.
    // Nothing here has a chart, so the numbers are zero — what matters is that
    // they ARRIVE, because `undefined` is what `describeMerge` reads as "this
    // outcome does not know" and says nothing about.
    expect(outcome.chartValues, "the merge said nothing about the chart values").toEqual({
      filled: 0,
      refused: 0,
      unreadable: 0,
      unplotted: 0,
    });
  });
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
  describe("when more slides arrive than the package held", () => {
    /**
     * `added` is measured from the DECK rather than from the plan, and that is
     * right: when the host lands fewer slides than it was handed, the deck knows
     * and the plan does not.
     *
     * It is wrong in the other direction, and the wrongness is not cosmetic.
     * `sweepPlan` refuses to sweep when the deck grew by more than the run added
     * — the clamp that keeps an undo off a stranger's slides — and an uncapped
     * `added` absorbs the excess, so `grew` and `added` are equal by construction
     * and that clamp can never fire. Six slides arriving across an insert of two
     * would have authorised deleting six.
     *
     * Capped, the same case leaves `grew > added` true at undo time and the sweep
     * refuses, which is the answer the rule was written to give.
     */
    it("counts only as many as it sent", async () => {
      host.readTemplate.mockResolvedValueOnce({ base64: await wholeDeck(), offset: 2 });
      host.insertDeck.mockResolvedValueOnce({ verdict: "yes", detail: "landed", landed: 6, before: 5, after: 11 });

      const out = await runMerge({ from: 3, to: 3, records: rows });
      const held = (await (await Pkg.open(host.insertDeck.mock.calls[0]?.[0] as string)).slidePaths()).length;

      expect(held, "the fixture stopped exercising the case").toBe(2);
      expect(out.added, "the deck's growth was taken as this run's own").toBe(held);
    });

    it("does not report a negative number of slides", async () => {
      // A deck that SHRANK across an insert. Nothing this run did, and nothing it
      // can take back.
      host.readTemplate.mockResolvedValueOnce({ base64: await wholeDeck(), offset: 2 });
      host.insertDeck.mockResolvedValueOnce({ verdict: "no", detail: "shrank", landed: -2, before: 5, after: 3 });

      const out = await runMerge({ from: 3, to: 3, records: rows });
      expect(out.added).toBe(0);
    });
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

  it("does not read an unaccountable deck as a torn merge", async () => {
    /**
     * A co-author or AutoSave landing a slide across the insert makes the deck
     * grow by MORE than the package held. `insertVerdict` grades that
     * `unknown` and writes a sentence naming the condition; this file branched
     * on whether anything landed and took the torn path, producing a sentence
     * that cannot be true — "PowerPoint took only part of the merge: all 3
     * row(s) landed complete" — and then telling the user to take the slides
     * back, which `sweepPlan` refuses for exactly this shape. The offer was a
     * dead end with a slide-deleting button on it.
     */
    host.readTemplate.mockResolvedValueOnce({ base64: await block(), offset: 0 });
    host.insertDeck.mockResolvedValueOnce({
      verdict: "unknown",
      detail: "the deck grew by 8 while the package held 6 slide(s), so this run cannot say which of them are its own",
      landed: 8,
      before: 2,
      after: 10,
    });

    const out = await runMerge({ from: 1, to: 2, records: rows });
    expect(out.ok).toBe(false);
    expect(out.detail, "the self-contradicting sentence").not.toMatch(/took only part/);
    expect(out.detail, "advice the undo refuses to carry out").not.toMatch(/take the slides back/i);
    expect(out.detail, "the verdict's own sentence, which names the condition").toContain(
      "cannot say which of them are its own",
    );
    expect(out.accountable, "the pane may not offer to sweep slides this run cannot identify").toBe(false);
  });

  it("does not read an over-grown deck as accountable just because the call also raised", async () => {
    /**
     * `insertVerdict` grades `unknown` only when the call did NOT raise, so an
     * insert that timed out AND over-grew — the case its own docstring
     * describes, "a call can raise and still have done the work" — came back
     * `threw`. Reading the verdict string rather than the deltas let both
     * halves of the defect through on that branch: the torn sentence, and an
     * undo card over slides `sweepPlan` would decline.
     *
     * Found by an adversarial review of the commit that was supposed to close
     * it, which is why the rule is on the numbers now.
     */
    host.readTemplate.mockResolvedValueOnce({ base64: await block(), offset: 0 });
    host.insertDeck.mockResolvedValueOnce({
      verdict: "threw",
      detail: "the call threw: GeneralException, and 8 slide(s) landed anyway",
      landed: 8,
      before: 2,
      after: 10,
    });

    const out = await runMerge({ from: 1, to: 2, records: rows });
    expect(out.accountable, "the pane may not offer to sweep slides this run cannot identify").toBe(false);
    expect(out.detail, "the self-contradicting sentence").not.toMatch(/took only part/);
    expect(out.detail).toContain("cannot say which of them are its own");
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

  it("reports the same set of facts whether the host took it or tore it", () => {
    /**
     * The two full outcomes differ in exactly two fields — `ok` and the
     * sentence — and everything else is what the RUN produced, which does not
     * change because the host misbehaved on the way out. They were written out
     * separately, sixteen identical fields each, and the file still carried two
     * comments explaining that `paragraphsMerged` and `pictures` are on the
     * success path "too": notes about an asymmetry somebody had already gone
     * back to close.
     *
     * Shared as one object now, so the type keeps them in step. This is the
     * behavioural half — a re-inlined pair that forgets a field on one side
     * fails here, where a compiler would be perfectly happy.
     *
     * What it does NOT catch, stated because it was checked: a field present
     * with the value `undefined`. `Object.keys` counts the key either way, so
     * this is a guard about the SHAPE of the two outcomes and not about their
     * values. Comparing values would fail on `detail` and on every count that
     * legitimately differs between a whole insert and a torn one.
     */
    return (async () => {
      host.readTemplate.mockResolvedValueOnce({ base64: await block(), offset: 0 });
      host.insertDeck.mockResolvedValueOnce({ verdict: "yes", detail: "landed", landed: 6, before: 2, after: 8 });
      const whole = await runMerge({ from: 1, to: 2, records: rows });

      host.readTemplate.mockResolvedValueOnce({ base64: await block(), offset: 0 });
      host.insertDeck.mockResolvedValueOnce({ verdict: "no", detail: "5 of 6", landed: 5, before: 2, after: 7 });
      const torn = await runMerge({ from: 1, to: 2, records: rows });

      // The premise: one of each. Without this the comparison below could be
      // between two outcomes of the same kind and would prove nothing.
      expect(whole.ok).toBe(true);
      expect(torn.ok).toBe(false);

      const keys = (o: object) => Object.keys(o).sort();
      expect(keys(torn)).toEqual(keys(whole));
      // And it is a real set, not an empty object agreeing with itself.
      expect(keys(whole).length).toBeGreaterThan(10);
    })();
  });
});

describe("a block with no fields on it yet", () => {
  /**
   * The seam the five-step order turns on, and the one place it must NOT be
   * uniform.
   *
   * The pane picks the slides that repeat before any `{{field}}` is on them —
   * it has to, because the names to type are the data's own column headers and
   * the data is attached at the step after. So a template read has to ANSWER
   * for an empty block. A merge must still refuse one: N identical copies is
   * never what anybody meant and is expensive to undo once it is in the deck.
   *
   * `prepareBlock` is the same function on both paths, so the difference is one
   * flag (`allowEmpty`) passed by exactly one caller. This is the test that
   * says the flag has not leaked onto the destructive path.
   */
  const empty = async (): Promise<void> => {
    const bytes = await makeDeck([{ paragraphs: [["Click to add title"]] }, { paragraphs: [["after"]] }]);
    host.readTemplate.mockResolvedValue({ base64: Buffer.from(bytes).toString("base64"), offset: 0 });
  };

  it("is an ANSWER to inspectBlock, with no fields in it", async () => {
    await empty();
    const report = await inspectBlock({ from: 1, to: 1 });
    expect(report.ok, report.ok ? "" : report.detail).toBe(true);
    expect(report.fields).toEqual([]);
  });

  it("is still a REFUSAL from runMerge, in the words that say what to type", async () => {
    await empty();
    const out = await runMerge({ from: 1, to: 1, records });
    expect(out.ok).toBe(false);
    expect(out.detail).toContain("{{fields}}");
    expect(out.added, "nothing may reach the deck").toBe(0);
    expect(host.insertDeck).not.toHaveBeenCalled();
  });
});
