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

describe("placeholders in a chart or SmartArt", () => {
  /**
   * They are ORDINARY fields now. Until picture fields shipped they were
   * reported as unfillable — correctly, and the pane said so — and this scan
   * kept them in a list of their own so the fields step would not ask for data
   * that could never be used.
   *
   * What matters is the property that outlived both behaviours: **this scan and
   * `runPlan` must read the same parts.** A scan that reads fewer refuses a
   * block it would have merged; a scan that reads more asks for a column
   * nothing will fill. It has been wrong in both directions — notes were
   * missing from the scan, charts were in it and not in the merge — so the
   * tests below are about the two lists agreeing.
   */
  it("counts a placeholder in a chart as a field", async () => {
    const deck = await makeDeck([
      { paragraphs: [["Hello {{First}}"]], chart: "Sales for {{Region}}" },
      { paragraphs: [["after"]] },
    ]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok).toBe(true);
    if (!prepared.ok) return;
    expect(prepared.fields).toEqual(["First", "Region"]);
  });

  it("counts one in a chart's category labels, which are not paragraphs", async () => {
    // The labels a user writes are `<c:v>` in a string cache. The merge fills
    // them, so the scan has to see them — or the pane offers no column for the
    // one field the template most obviously has.
    const deck = await makeDeck([
      { paragraphs: [["Hello {{First}}"]], chart: { categories: ["{{Region}}", "Other"] } },
      { paragraphs: [["after"]] },
    ]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok && prepared.fields).toEqual(["First", "Region"]);
  });

  it("counts one in SmartArt", async () => {
    const deck = await makeDeck([
      { paragraphs: [["Hello"]], smartArt: ["{{Region}} team"] },
      { paragraphs: [["after"]] },
    ]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok && prepared.fields).toEqual(["Region"]);
  });

  it("accepts a block whose ONLY placeholder is in a chart", async () => {
    // This was a refusal — "move the text onto the slide itself" — and the
    // merge it refused is the merge that now runs. A stale refusal is worse
    // than none: it sends the author to undo the thing that works.
    const deck = await makeDeck([{ paragraphs: [["a title"]], chart: "{{Region}}" }, { paragraphs: [["after"]] }]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok).toBe(true);
    expect(prepared.ok && prepared.fields).toEqual(["Region"]);
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
    expect(prepared.ok && prepared.fields).toEqual(["OnTheSlide", "LongFieldName"]);
  });

  it("ignores a chart on a slide outside the block", async () => {
    /**
     * The reason this reads the parts THIS slide relates to rather than the
     * package at large. On the route below API 1.10 the template comes back as
     * the WHOLE deck, so a package-wide scan would name a chart on slide 40 and
     * ask for a column the block has no use for.
     */
    const deck = await makeDeck([
      { paragraphs: [["{{First}}"]] },
      { paragraphs: [["elsewhere"]], chart: "{{NotMine}}" },
    ]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok && prepared.fields).toEqual(["First"]);
  });
});

describe("the word PowerPoint has already taken", () => {
  it("names the syntax rather than saying 'no placeholders'", async () => {
    /**
     * PowerPoint calls its own empty content boxes placeholders — "Click to add
     * title" IS a placeholder in its vocabulary. So the old refusal told a user
     * staring at two of them that the slide had none, on first contact with the
     * add-in, on a fresh deck. It reads as the thing being broken, and it was
     * reported as exactly that.
     *
     * The refusal has to name what to TYPE. This asserts both halves: the
     * syntax is there, and the bare word that collides is not.
     */
    const deck = await makeDeck([{ paragraphs: [["Click to add title"]] }, { paragraphs: [["after"]] }]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok).toBe(false);
    if (prepared.ok) return;
    expect(prepared.why).toContain("{{fields}}");
    // Where the names COME FROM, not an invented one: the user has not attached
    // data at this step and does not know their column headers yet, which is
    // the objection this sentence was rewritten to answer.
    expect(prepared.why, "tells the user what to type").toContain("column headers");
    expect(prepared.why).toContain("{{First}}");
    expect(prepared.why, "PowerPoint's word for its own empty boxes").not.toMatch(/\bno placeholders\b/);
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
    expect(prepared.ok || prepared.why).toContain("no {{fields}}");
  });
});
