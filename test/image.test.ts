import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { readImage } from "../src/core/image/read.js";
import { NO_INSET, containFillRect, coverSrcRect } from "../src/core/image/fill.js";

/**
 * Real bytes, not hand-built headers.
 *
 * A header assembled in the test is a header written to the same understanding
 * as the reader, so the two agree by construction and neither is checked. These
 * two files are rendered by Chromium and committed, so the reader is measured
 * against what an encoder actually produces.
 */
const PNG = readFileSync("test/fixtures/wide.png"); // 64 x 32
const JPEG = readFileSync("test/fixtures/tall.jpg"); // 30 x 90

describe("reading an image's own header", () => {
  it("reads a real PNG", () => {
    expect(readImage(PNG)).toEqual({
      kind: "png",
      contentType: "image/png",
      extension: "png",
      width: 64,
      height: 32,
    });
  });

  it("reads a real JPEG, whose size is not at a fixed offset", () => {
    // A JPEG is a chain of segments and the size lives in whichever
    // start-of-frame marker the encoder used. This one has to be walked to.
    expect(readImage(JPEG)).toMatchObject({ kind: "jpeg", width: 30, height: 90 });
  });

  it("does not take the format from the FILE NAME", () => {
    // A `.jpg` that is really a PNG is an ordinary thing to find in a folder of
    // exports, and declaring the wrong content type is how a deck opens as
    // damaged. Nothing here has ever seen the name.
    expect(readImage(PNG)?.contentType).toBe("image/png");
  });

  it("answers undefined for anything it cannot read, and never throws", () => {
    for (const bad of [
      new Uint8Array(),
      new Uint8Array([1, 2, 3]),
      new Uint8Array(64), // all zeroes
      PNG.subarray(0, 20), // a PNG truncated inside IHDR
      new Uint8Array([0xff, 0xd8, 0xff]), // a JPEG that stops at its signature
    ]) {
      expect(() => readImage(bad)).not.toThrow();
      expect(readImage(bad)).toBeUndefined();
    }
  });

  it("refuses a PNG whose first chunk is not IHDR", () => {
    // The size is read from a fixed offset that is only correct because IHDR is
    // required to come first. Checked rather than assumed.
    const bent = new Uint8Array(PNG);
    bent[12] = 0x49;
    bent[13] = 0x45; // "IE" — not IHDR
    expect(readImage(bent)).toBeUndefined();
  });

  it("reads a GIF, which stores its size little-endian", () => {
    const gif = new Uint8Array(20);
    gif.set([0x47, 0x49, 0x46, 0x38, 0x39, 0x61]);
    gif.set([0x40, 0x00, 0x20, 0x00], 6); // 64 x 32
    expect(readImage(gif)).toMatchObject({ kind: "gif", width: 64, height: 32 });
  });

  it("reads a top-down BMP, whose height is NEGATIVE", () => {
    // The sign is a direction, not a size. Taken raw it is a negative extent
    // and every ratio built from it comes out backwards.
    const bmp = new Uint8Array(30);
    bmp.set([0x42, 0x4d]);
    new DataView(bmp.buffer).setInt32(18, 64, true);
    new DataView(bmp.buffer).setInt32(22, -32, true);
    expect(readImage(bmp)).toMatchObject({ kind: "bmp", width: 64, height: 32 });
  });
});

describe("covering a shape without distorting", () => {
  it("trims nothing when the ratios already match", () => {
    expect(coverSrcRect({ w: 200, h: 100 }, { w: 64, h: 32 })).toEqual(NO_INSET);
  });

  it("trims the SIDES when the image is proportionally wider", () => {
    // A 2:1 image in a 1:1 box: half the width goes, a quarter off each side.
    expect(coverSrcRect({ w: 100, h: 100 }, { w: 64, h: 32 })).toEqual({ l: 25000, t: 0, r: 25000, b: 0 });
  });

  it("trims TOP AND BOTTOM when the image is proportionally taller", () => {
    expect(coverSrcRect({ w: 100, h: 100 }, { w: 30, h: 90 })).toEqual({ l: 0, t: 33334, r: 0, b: 33333 });
  });

  it("splits an odd trim so the two sides still sum to the whole", () => {
    // Rounding each side on its own drifts, and the shape ends up a hairline
    // uncovered down one edge — visible on a photo against a coloured
    // background, and never visible in a unit that checks one side.
    const inset = coverSrcRect({ w: 100, h: 100 }, { w: 30, h: 90 });
    expect(inset.t + inset.b).toBe(66667);
  });

  it("answers NO_INSET rather than dividing by zero", () => {
    for (const [shape, image] of [
      [
        { w: 0, h: 100 },
        { w: 64, h: 32 },
      ],
      [
        { w: 100, h: 0 },
        { w: 64, h: 32 },
      ],
      [
        { w: 100, h: 100 },
        { w: 0, h: 32 },
      ],
      [
        { w: 100, h: 100 },
        { w: 64, h: 0 },
      ],
    ] as const) {
      expect(coverSrcRect(shape, image)).toEqual(NO_INSET);
    }
  });

  it("never trims both axes, because one of them always fits exactly", () => {
    for (const image of [
      { w: 1, h: 1000 },
      { w: 1000, h: 1 },
      { w: 3, h: 7 },
      { w: 1920, h: 1080 },
    ]) {
      const inset = coverSrcRect({ w: 400, h: 300 }, image);
      expect(inset.l + inset.r === 0 || inset.t + inset.b === 0, JSON.stringify(image)).toBe(true);
    }
  });

  it("never trims away the whole image", () => {
    for (const image of [
      { w: 1, h: 4000 },
      { w: 4000, h: 1 },
    ]) {
      const inset = coverSrcRect({ w: 400, h: 300 }, image);
      expect(inset.l + inset.r, JSON.stringify(image)).toBeLessThan(100000);
      expect(inset.t + inset.b, JSON.stringify(image)).toBeLessThan(100000);
    }
  });
});

describe("containing an image inside a shape", () => {
  it("insets nothing when the ratios already match", () => {
    expect(containFillRect({ w: 200, h: 100 }, { w: 64, h: 32 })).toEqual(NO_INSET);
  });

  it("puts the bars ABOVE AND BELOW a proportionally wider image", () => {
    // The mirror of cover, and the axis is the opposite one: a wide image
    // spans the width and leaves space top and bottom. Getting this backwards
    // squeezes the image into a strip instead of letterboxing it.
    expect(containFillRect({ w: 100, h: 100 }, { w: 64, h: 32 })).toEqual({ l: 0, t: 25000, r: 0, b: 25000 });
  });

  it("puts the bars LEFT AND RIGHT of a proportionally taller image", () => {
    expect(containFillRect({ w: 100, h: 100 }, { w: 30, h: 90 })).toEqual({ l: 33334, t: 0, r: 33333, b: 0 });
  });

  it("is the opposite axis from cover, for the same pair", () => {
    // Stated as a property because the two functions are a paragraph apart and
    // look alike enough to be edited into each other.
    const shape = { w: 400, h: 300 };
    for (const image of [
      { w: 1920, h: 1080 },
      { w: 600, h: 900 },
    ]) {
      const cropped = coverSrcRect(shape, image);
      const inset = containFillRect(shape, image);
      expect(cropped.l + cropped.r > 0, JSON.stringify(image)).toBe(inset.t + inset.b > 0);
      expect(cropped.t + cropped.b > 0, JSON.stringify(image)).toBe(inset.l + inset.r > 0);
    }
  });
});
