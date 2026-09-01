import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";
import { prepareBlock } from "../src/core/merge/prepare.js";
import { Pkg } from "../src/core/pptx/pkg.js";
import { makeDeck } from "./fixtures/deck.js";
import { C_NS, elements } from "../src/core/pptx/xml.js";

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
    // A positive offset proves the whole file came back, so "the deck" is what
    // came back and the number is one the user can count.
    expect(out.why).toContain("the deck that came back");
  });

  it("does not call an exported block 'the deck' when only the block came back", async () => {
    // On the subset route PowerPoint sends back the exported BLOCK, so the
    // count in this sentence is the size of that export — "the deck that came
    // back has 3" to somebody looking at a deck of thirty. `offsetInPackage` is
    // zero there, which is what tells the two routes apart.
    const pkg = await deck(3);
    const out = await prepareBlock(pkg, { from: 1, to: 9, offsetInPackage: 0 }, "r1");
    expect(out.ok).toBe(false);
    if (out.ok) return;
    expect(out.why, "a deck the user cannot count").not.toContain("the deck that came back");
    expect(out.why).toContain("the slides PowerPoint sent back");
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

  it("inflates each chart workbook once, not once per reader", async () => {
    /**
     * Two readers walk the same workbook on this path — the value cells and the
     * text — and each used to inflate it for itself, doubling the cost of the
     * step on any deck with charts. They share one inflate now.
     *
     * Safe only because both are dry runs whose resolver writes nothing: both
     * readers MUTATE the zip when they fill something, so a shared book on the
     * merge path would let one see the other's edits.
     */
    const deck = await makeDeck([
      { paragraphs: [["Hello {{First}}"]], chart: { title: "Sales", workbook: ["{{Region}}"] } },
      { paragraphs: [["after"]] },
    ]);
    const pkg = await Pkg.open(deck);
    const load = vi.spyOn(JSZip, "loadAsync");
    try {
      const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
      expect(prepared.ok).toBe(true);
      const workbooks = load.mock.calls.length;
      expect(workbooks, "the chart's workbook is opened at all").toBeGreaterThan(0);
      expect(workbooks, "and only once, by both readers together").toBe(1);
    } finally {
      load.mockRestore();
    }
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

  it("names a picture field written where no picture can go", async () => {
    /**
     * `placeImages` fills a SHAPE on a slide. A `{{Photo|image}}` on the notes
     * page is filled by nothing and printed as written, so the raw placeholder
     * reaches presenter view and every printed handout — and if it is the
     * block's only picture field, the pane never even offers the file picker,
     * because `imageFields` is what turns that on.
     *
     * Filling it is not on the table; saying so before the merge is.
     */
    const deck = await makeDeck([
      { paragraphs: [["Quarterly review"]], notes: "Bring {{Photo|image}} to the meeting" },
      { paragraphs: [["after"]] },
    ]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok, prepared.ok ? "" : prepared.why).toBe(true);
    if (!prepared.ok) return;
    expect(prepared.imageFieldsOffSlide, "a field nothing will fill, said out loud").toEqual(["Photo"]);
    // And it is NOT offered as one the picker can serve, which would promise a
    // placement that cannot happen.
    expect(prepared.imageFields).toEqual([]);
  });

  it("does not call a picture field on a SLIDE off-slide", async () => {
    const deck = await makeDeck([{ paragraphs: [["{{Photo|image}}"]] }, { paragraphs: [["after"]] }]);
    const pkg = await Pkg.open(deck);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "run1");
    expect(prepared.ok).toBe(true);
    if (!prepared.ok) return;
    expect(prepared.imageFields).toEqual(["Photo"]);
    expect(prepared.imageFieldsOffSlide).toEqual([]);
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

describe("a block whose only field is the number its chart plots", () => {
  /**
   * `prepare.ts` states the rule this breaks: "this list and `runPlan`'s are
   * the same list". A scan that reads fewer parts than the merge writes REFUSES
   * a block it would have merged, telling the author to go and type field names
   * onto a slide that already carries one.
   *
   * It had happened twice — speaker notes, then chart labels — and both are
   * described in that file. This was the third: a value cell holds its
   * placeholder in the WORKBOOK, and `fieldsIn` cannot see it, because
   * `<c:numCache>` is left out of the text pass on purpose (a formatted number
   * is unplottable). So the workbook was not read at all, on a rationale that
   * was true until chart numbers became a feature.
   *
   * The scan is now a dry run of `mergeChartNumbers` itself, so the two cannot
   * hold different opinions about which cells carry a placeholder.
   */
  const CHART = {
    paragraphs: [["Quarterly revenue"]],
    chart: {
      title: "Revenue",
      categories: ["North", "South"],
      workbook: ["North"],
      values: ["{{Revenue}}", "42"],
    },
  };

  async function chartDeck(): Promise<Pkg> {
    return Pkg.open(await makeDeck([CHART, { paragraphs: [["after"]] }]));
  }

  it("is accepted, and the field is named", async () => {
    const prepared = await prepareBlock(await chartDeck(), { from: 1, to: 1, offsetInPackage: 0 }, "n");
    expect(prepared.ok || prepared.why, "refused a block the merge would have filled").toBe(true);
    expect(prepared.ok && prepared.fields).toEqual(["Revenue"]);
  });

  it("and the scan that finds it writes nothing", async () => {
    /**
     * The dry run drives the real merge with a resolver that answers null, so
     * "it does not write" is a claim about a code path, not a design. Measured
     * at the only two places that path can write: the embedded workbook, which
     * `mergeChartNumbers` repacks with `setBytes`, and the chart's cached
     * values, which it edits in the document.
     */
    const pkg = await chartDeck();
    const embedding = pkg.partNames().find((p) => p.endsWith(".xlsx")) as string;
    const before = Buffer.from(await pkg.bytes(embedding));
    const cached = async (): Promise<(string | null)[]> =>
      elements(await pkg.doc("ppt/charts/chart1.xml"), C_NS, "numCache")
        .flatMap((c) => elements(c, C_NS, "v"))
        .map((v) => v.textContent);

    const cacheBefore = await cached();
    await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "n");

    expect(Buffer.compare(before, Buffer.from(await pkg.bytes(embedding))), "the workbook was repacked").toBe(0);
    expect(await cached(), "the cached values moved").toEqual(cacheBefore);
    expect(cacheBefore, "the fixture stopped carrying a placeholder").toContain("{{Revenue}}");
  });
});

describe("which fields ask for a picture", () => {
  /**
   * The pane's picker is shown when the data refers to pictures, and it decided
   * that from the DATA's detected types alone. The engine decides from the
   * FIELD's format, so the two disagreed about any column the type detector had
   * turned down — see `pictureColumns` in the pane.
   *
   * `prepare` reads the slides already; naming the picture fields costs nothing
   * and is the only place that can answer it.
   */
  it("names them, apart from the ordinary fields", async () => {
    const pkg = await Pkg.open(
      await makeDeck([
        { paragraphs: [["{{Name}}"], ["{{Photo|image}}"], ["{{Logo|image-fit}}"]] },
        { paragraphs: [["after"]] },
      ]),
    );
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "n");
    expect(prepared.ok).toBe(true);
    if (!prepared.ok) return;
    expect(prepared.fields).toEqual(["Name", "Photo", "Logo"]);
    expect(prepared.imageFields).toEqual(["Photo", "Logo"]);
  });

  it("and answers an empty list when nothing asks for one", async () => {
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["{{Name}}"]] }, { paragraphs: [["after"]] }]));
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "n");
    expect(prepared.ok && prepared.imageFields).toEqual([]);
  });
});
