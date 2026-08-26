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
