import { describe, expect, it } from "vitest";
import { prepareBlock } from "../src/core/merge/prepare.js";
import { Pkg } from "../src/core/pptx/pkg.js";
import { makeDeck } from "./fixtures/deck.js";

/** A deck of `n` slides, the middle ones carrying placeholders. */
async function deck(n: number, withFields = true): Promise<Pkg> {
  return Pkg.open(
    await makeDeck(
      Array.from({ length: n }, (_, i) => ({
        paragraphs: [[withFields && i > 0 ? `Slide ${i + 1} for {{Name}}` : `Slide ${i + 1}`]],
      })),
    ),
  );
}

describe("turning slide numbers into a block", () => {
  it("takes the slides the user pointed at, in order", async () => {
    const pkg = await deck(6);
    const out = await prepareBlock(pkg, { from: 2, to: 4, offsetInPackage: 1 }, "r1");
    expect(out.ok).toBe(true);
    if (!out.ok) return;
    expect(out.block.slides.map((s) => s.seq)).toEqual([1, 2, 3]);
    expect(out.block.slides.map((s) => s.path)).toEqual([
      "ppt/slides/slide2.xml",
      "ppt/slides/slide3.xml",
      "ppt/slides/slide4.xml",
    ]);
    expect(out.fields).toEqual(["Name"]);
  });

  it("reads the offset, because the two read routes return different packages", async () => {
    // A subset export holds ONLY the template, so the block starts at zero; a
    // whole-deck read holds everything and it starts where it sat in the deck.
    // A caller that assumes either merges the wrong slides, and the output
    // looks deliberate.
    const pkg = await deck(6);
    const whole = await prepareBlock(pkg, { from: 4, to: 4, offsetInPackage: 3 }, "r1");
    const subset = await prepareBlock(pkg, { from: 4, to: 4, offsetInPackage: 1 }, "r1");
    expect(whole.ok && whole.block.slides[0]?.path).toBe("ppt/slides/slide4.xml");
    expect(subset.ok && subset.block.slides[0]?.path).toBe("ppt/slides/slide2.xml");
  });

  it("refuses a block that runs off the end of the deck it was given", async () => {
    const pkg = await deck(3);
    const out = await prepareBlock(pkg, { from: 2, to: 9, offsetInPackage: 1 }, "r1");
    expect(out.ok).toBe(false);
    if (out.ok) return;
    // Says both numbers, so the user can see which one is wrong.
    expect(out.why).toContain("2");
    expect(out.why).toContain("3");
  });

  it("refuses a block that ends before it starts", async () => {
    const pkg = await deck(6);
    const out = await prepareBlock(pkg, { from: 5, to: 2, offsetInPackage: 4 }, "r1");
    expect(out.ok).toBe(false);
    if (!out.ok) expect(out.why).toContain("ends before it starts");
  });

  it("refuses slide 0, which is not a slide anybody can see", async () => {
    const pkg = await deck(6);
    const out = await prepareBlock(pkg, { from: 0, to: 2, offsetInPackage: 0 }, "r1");
    expect(out.ok).toBe(false);
    if (!out.ok) expect(out.why).toContain("numbered from 1");
  });

  it("refuses a block with no placeholders rather than making N identical copies", async () => {
    // The engine cannot see this as an error — it would clone happily. It is
    // never what anybody meant, and it is expensive to undo once it is in the
    // deck.
    const pkg = await deck(3, false);
    const out = await prepareBlock(pkg, { from: 1, to: 2, offsetInPackage: 0 }, "r1");
    expect(out.ok).toBe(false);
    if (!out.ok) expect(out.why).toContain("identical");
  });

  it("carries a condition through to the slide it belongs to", async () => {
    const pkg = await deck(6);
    const out = await prepareBlock(pkg, { from: 2, to: 4, offsetInPackage: 1, conditions: { 3: "HasBonus" } }, "r1");
    expect(out.ok).toBe(true);
    if (!out.ok) return;
    // Keyed by SLIDE NUMBER, not by position in the block: those differ for
    // every block that does not start at slide 1, and the pane speaks numbers.
    expect(out.block.slides.map((s) => s.condition)).toEqual([undefined, "HasBonus", undefined]);
  });
});

describe("placeholders the engine does not reach", () => {
  /**
   * A chart's labels live in `ppt/charts/chartN.xml` with an embedded workbook
   * behind them; SmartArt's live in `ppt/diagrams/dataN.xml`. Neither is a
   * `<a:p>` on the slide, so `mergeDocument` never touches them and `fieldsIn`
   * never reports them.
   *
   * Not merging them is a stated limit. Not SAYING so was the defect: the
   * author puts `{{Region}}` in a chart title, the pane counts the placeholders
   * it can see, and 240 slides ship with the braces on them.
   */
  it("reports a placeholder in a chart instead of passing over it", async () => {
    const deck = await makeDeck([
      { paragraphs: [["Hello {{First}}"]], chart: "Sales for {{Region}}" },
      { paragraphs: [["after"]] },
    ]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok).toBe(true);
    if (!prepared.ok) return;
    // Kept APART from the fields, not folded in: it is not a candidate for a
    // column, so counting it would make the fields step ask for data that
    // could never be used.
    expect(prepared.fields).toEqual(["First"]);
    expect(prepared.unmergeable).toEqual(["Region"]);
  });

  it("does not say a block has no placeholders when one is in its chart", async () => {
    /**
     * "No placeholders" is true — the engine cannot fill a chart — and useless:
     * the author placed one, is looking at it, and is being told it is not
     * there. The complaint this pass exists for, arriving on the one path
     * where it reads as the engine being broken.
     */
    const deck = await makeDeck([{ paragraphs: [["a title"]], chart: "{{Region}}" }, { paragraphs: [["after"]] }]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok).toBe(false);
    if (prepared.ok) return;
    expect(prepared.why).toContain("Region");
    expect(prepared.why).toContain("chart or SmartArt");
    expect(prepared.why, "the unhelpful sentence").not.toContain("no placeholders");
  });

  it("finds one the host has split across runs", async () => {
    // The fixture splits its chart text in half deliberately. A placeholder
    // split across two runs is the ordinary state of one after an edit, and it
    // is what a regex over the raw markup would miss while reporting the tidy
    // ones — the same reason `mergeParagraph` matches against joined text.
    const deck = await makeDeck([
      { paragraphs: [["{{OnTheSlide}}"]], chart: "{{LongFieldName}}" },
      { paragraphs: [["y"]] },
    ]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok && prepared.unmergeable).toEqual(["LongFieldName"]);
  });

  it("says nothing when no chart holds a placeholder", async () => {
    const deck = await makeDeck([
      { paragraphs: [["{{First}}"]], chart: "Sales by quarter" },
      { paragraphs: [["after"]] },
    ]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok && prepared.unmergeable).toEqual([]);
  });

  it("ignores a chart on a slide outside the block", async () => {
    /**
     * The reason this reads the parts THIS slide relates to rather than the
     * package at large. On the route below API 1.10 the template comes back as
     * the WHOLE deck, so a package-wide scan would name a chart on slide 40 and
     * send the user hunting through a template that is fine.
     */
    const deck = await makeDeck([
      { paragraphs: [["{{First}}"]] },
      { paragraphs: [["elsewhere"]], chart: "{{NotMine}}" },
    ]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok && prepared.unmergeable).toEqual([]);
  });
});

describe("placeholders in the speaker notes", () => {
  /**
   * `runPlan` merges the notes page and always has — a template whose notes
   * read "Call {{Name}} afterwards" otherwise ships that verbatim on every
   * handout and in the presenter view. This scan only ever read the SLIDE.
   *
   * So a block whose placeholders live in the notes was refused with "no
   * placeholders, so every copy would be identical" — a sentence about a merge
   * that would have filled them. The mirror of the chart case, and the worse
   * direction: there the pane reported fields it cannot merge; here it hid
   * fields it can and blocked the merge on the strength of it.
   */
  it("counts a field the merge would fill", async () => {
    const deck = await makeDeck([
      { paragraphs: [["Quarterly review"]], notes: "Call {{Name}} afterwards" },
      { paragraphs: [["after"]] },
    ]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok, prepared.ok ? "" : prepared.why).toBe(true);
    if (!prepared.ok) return;
    expect(prepared.fields).toEqual(["Name"]);
  });

  it("merges the slide's and the notes' fields into one list, without duplicates", async () => {
    const deck = await makeDeck([
      { paragraphs: [["Hello {{First}}"]], notes: "Ring {{First}} about {{Topic}}" },
      { paragraphs: [["after"]] },
    ]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok && prepared.fields).toEqual(["First", "Topic"]);
  });

  it("still refuses a block with nothing anywhere", async () => {
    // The refusal has to survive: a merge with no placeholders at all produces
    // N identical copies, which is expensive to undo once it is in the deck.
    const deck = await makeDeck([
      { paragraphs: [["Quarterly review"]], notes: "no fields here" },
      { paragraphs: [["after"]] },
    ]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok).toBe(false);
    expect(prepared.ok || prepared.why).toContain("no placeholders");
  });
});
