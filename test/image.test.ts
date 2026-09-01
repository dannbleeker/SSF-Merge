import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { readImage } from "../src/core/image/read.js";
import { NO_INSET, containFillRect, coverSrcRect } from "../src/core/image/fill.js";
import { shapeBox } from "../src/core/image/place.js";
import { A_NS, P_NS, elements, parseXml } from "../src/core/pptx/xml.js";

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
    // The DIB header states its own length at 14, and this one went in as zero
    // — not a BMP any encoder writes, and it passed only while the reader
    // ignored that field. Reading it is what tells a 12-byte OS/2 header from
    // a 40-byte one, so the fixture has to say which it is. The claim below is
    // untouched.
    new DataView(bmp.buffer).setInt32(14, 40, true);
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
      // CONTAIN carries the identical guard and nothing was checking it: the
      // test's name is about dividing by zero, and half the functions that can
      // were outside the loop. Found by `scripts/mutate-core.mjs` — loosening
      // contain's guard to `< 0` left the suite green, and a zero-sided shape
      // then produces a 100% inset rather than none.
      expect(containFillRect(shape, image)).toEqual(NO_INSET);
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

describe("header shapes a real encoder writes and this reader did not expect", () => {
  /**
   * Hand-built, against the argument at the top of this file — and framed to
   * survive it.
   *
   * The objection to a hand-built header is that it is written to the same
   * understanding as the reader, so the two agree by construction. These do not
   * assert an absolute size read out of bytes I wrote. Each asserts that a
   * LEGAL VARIATION does not change the answer, against a control in the shape
   * an ordinary encoder produces — and the walker that reads the control is the
   * same one the two committed files exercise above. If my understanding of the
   * variation were wrong, the pair would disagree with each other.
   */
  const SOF0 = [0xc0, 0x00, 0x11, 0x08, 0x00, 0x64, 0x00, 0xc8, 0x03, 1, 0x11, 0, 2, 0x11, 0, 3, 0x11, 0];
  const jpeg = (...lead: number[]) => Uint8Array.from([0xff, 0xd8, ...lead, 0xff, ...SOF0]);

  it("reads a JPEG whose marker is padded with fill bytes", () => {
    /**
     * Any number of 0xFF bytes may pad the space before a marker. Read as a
     * marker itself, `FF FF C0 ...` takes the frame header's own first bytes
     * as a segment length and skips a nonsense distance — so a good JPEG came
     * back undefined, which the pane reports as an unreadable file.
     */
    const control = readImage(jpeg());
    expect(control?.width, "the control is not being read").toBe(200);
    expect(readImage(jpeg(0xff)), "one fill byte lost the frame").toEqual(control);
    expect(readImage(jpeg(0xff, 0xff, 0xff, 0xff)), "four fill bytes lost the frame").toEqual(control);
  });

  it("walks past a standalone marker that carries no length", () => {
    /**
     * `TEM` (`FF 01`) and the eight restart markers (`FF D0`-`FF D7`) are the
     * markers with no length after them. Skipped by a length they do not have,
     * the walk reads the FRAME HEADER's own first bytes as a segment length and
     * jumps a nonsense distance — the same failure as the fill-byte case above,
     * and the same symptom: a perfectly good JPEG reported as a file this
     * add-in cannot read, with the picture placeholder left standing.
     *
     * Same method as the test above: not an absolute size read out of bytes I
     * wrote, but a legal variation that must not change the control's answer.
     */
    const control = readImage(jpeg());
    expect(control?.width, "the control is not being read").toBe(200);
    expect(readImage(jpeg(0xff, 0x01)), "TEM").toEqual(control);
    expect(readImage(jpeg(0xff, 0xd0)), "RST0").toEqual(control);
    expect(readImage(jpeg(0xff, 0xd7)), "RST7").toEqual(control);
  });

  it("walks past padding that is not a fill byte", () => {
    /**
     * The walker advances one byte at a time until it finds `0xFF`, rather than
     * assuming the byte after a segment begins a marker. A stray byte between
     * segments is not something an encoder should write and is something a
     * reader meets; taken as the start of a marker it reads the byte after it
     * as a marker number and the two after that as a length.
     *
     * The padding has to come AFTER a segment, not straight after the SOI:
     * `readImage` requires `FF D8 FF` before it will call this a JPEG at all,
     * which is right — a marker must follow the start-of-image — and it is why
     * the first version of this test asserted a file the reader correctly
     * refuses.
     */
    const control = readImage(jpeg());
    // An APP0 of four bytes, then one stray byte before the frame header.
    const app0 = [0xff, 0xe0, 0x00, 0x04, 0x01, 0x02];
    expect(readImage(jpeg(...app0, 0x00)), "one stray byte").toEqual(control);
    expect(readImage(jpeg(...app0, 0x00, 0x13, 0x7f)), "three stray bytes").toEqual(control);
  });

  it("reads the same size from either BMP header", () => {
    /**
     * The 12-byte OS/2 header keeps width and height as two 16-bit numbers at
     * 18 and 20; every later header keeps them as two 32-bit numbers at 18 and
     * 22. Reading the second shape out of the first does not fail — it returns
     * `200 | (100 << 16)`, so a 200 x 100 bitmap measured 6553800 x 1572865 and
     * was cropped to a ratio with nothing to do with it.
     *
     * Nothing said the size was invented, which is what makes this worse than
     * the JPEG above: that one refused a good file, this one accepted a wrong
     * answer.
     */
    const head = [0x42, 0x4d, 0, 0, 0, 0, 0, 0, 0, 0, 26, 0, 0, 0];
    const os2 = Uint8Array.from([...head, 12, 0, 0, 0, 0xc8, 0x00, 0x64, 0x00, 1, 0, 24, 0]);
    const info = Uint8Array.from([...head, 40, 0, 0, 0, 0xc8, 0, 0, 0, 0x64, 0, 0, 0, 1, 0, 24, 0]);

    expect(readImage(info)?.width, "the control is not being read").toBe(200);
    expect(readImage(os2), "the two headers describe the same bitmap").toEqual(readImage(info));
  });

  it("refuses a BMP header length it does not know rather than guessing", () => {
    // The whole point of the branch. An unknown header is a size at an unknown
    // offset, and a wrong size is placed without complaint.
    const odd = Uint8Array.from([
      0x42, 0x4d, 0, 0, 0, 0, 0, 0, 0, 0, 26, 0, 0, 0, 14, 0, 0, 0, 0xc8, 0, 0x64, 0, 1, 0, 24, 0,
    ]);
    expect(readImage(odd)).toBeUndefined();
  });

  it("reads EVERY frame marker, not just the baseline one", () => {
    /**
     * `SOF0` is baseline and `SOF2` is PROGRESSIVE, which is what "save for
     * web" produces — so a reader that knew only baseline would call an
     * ordinary photo unreadable, leave the placeholder on the slide, and be
     * right about nothing.
     *
     * All thirteen are read correctly today. Nothing held it: replacing the
     * whole predicate with `marker === 0xc0` left the entire suite green,
     * because every fixture here and the one committed .jpg are baseline. A
     * behaviour no test can lose is a behaviour no test is keeping.
     *
     * The three gaps in the range are the point of listing them: `C4` is a
     * Huffman table, `C8` is reserved and `CC` is an arithmetic-coding table.
     * None is a frame, and a reader that took the whole `C0`-`CF` block would
     * read two arbitrary bytes of a Huffman table as a width.
     */
    const framed = (marker: number) => Uint8Array.from([0xff, 0xd8, 0xff, marker, ...SOF0.slice(1)]);
    for (const marker of [0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf]) {
      const read = readImage(framed(marker));
      expect(read?.width, `SOF at 0x${marker.toString(16)}`).toBe(200);
      expect(read?.height, `SOF at 0x${marker.toString(16)}`).toBe(100);
    }
    for (const marker of [0xc4, 0xc8, 0xcc]) {
      expect(readImage(framed(marker)), `0x${marker.toString(16)} is not a frame`).toBeUndefined();
    }
  });

  it("still walks past the segments that are not frame headers", () => {
    // Unchanged behaviour, asserted because the fill-byte branch sits directly
    // above these and an early `continue` would swallow them.
    expect(readImage(jpeg(0xff, 0xe1, 0x00, 0x08, 1, 2, 3, 4, 5, 6))?.width, "APP1/EXIF").toBe(200);
    expect(readImage(jpeg(0xff, 0xc4, 0x00, 0x04, 1, 2))?.width, "DHT").toBe(200);
    expect(readImage(Uint8Array.from([0xff, 0xd8, 0xff, 0xc0, 0x00, 0x11, 0x08, 0x00])), "truncated").toBeUndefined();
  });
});

describe("the geometry keeps the ratios it claims", () => {
  /**
   * The cases above are points somebody chose. These are the PROPERTIES, over
   * every combination of a dozen sizes in both dimensions of both boxes —
   * 20,736 of them — and each is stated in terms of what the mode promises
   * rather than in terms of the arithmetic that produces it:
   *
   * - COVER scales until the image covers the box and trims the overflow, so
   *   the visible SLICE of the image must have the SHAPE's aspect ratio.
   * - CONTAIN scales until the whole image fits and centres it, so the rect it
   *   is placed in must have the IMAGE's aspect ratio.
   *
   * Neither restates `coverSrcRect`. Both can be worked out from the picture on
   * the slide, which is the point: a formula checked against itself proves
   * nothing, and this is the arithmetic behind every merged photo.
   */
  const M = 100000;
  const SIZES = [1, 2, 3, 5, 7, 16, 30, 64, 100, 333, 1000, 1920];

  /** How far apart the two aspect ratios are. */
  const mismatch = (shape: { w: number; h: number }, image: { w: number; h: number }): number =>
    Math.max(shape.w / shape.h / (image.w / image.h), image.w / image.h / (shape.w / shape.h));

  it("holds for every shape and picture a deck could plausibly hold", () => {
    let worstCover = 0;
    let worstContain = 0;
    let checked = 0;

    for (const sw of SIZES)
      for (const sh of SIZES)
        for (const iw of SIZES)
          for (const ih of SIZES) {
            const shape = { w: sw, h: sh };
            const image = { w: iw, h: ih };
            // Beyond about a thousandfold the unit itself is the limit — see
            // the next test, which pins that boundary rather than hiding it.
            if (mismatch(shape, image) > 20) continue;
            checked++;

            const c = coverSrcRect(shape, image);
            const visW = image.w * (1 - (c.l + c.r) / M);
            const visH = image.h * (1 - (c.t + c.b) / M);
            worstCover = Math.max(worstCover, Math.abs(visW / visH - shape.w / shape.h) / (shape.w / shape.h));

            const f = containFillRect(shape, image);
            const boxW = shape.w * (1 - (f.l + f.r) / M);
            const boxH = shape.h * (1 - (f.t + f.b) / M);
            worstContain = Math.max(worstContain, Math.abs(boxW / boxH - image.w / image.h) / (image.w / image.h));
          }

    expect(checked, "the sweep stopped covering anything").toBeGreaterThan(5000);
    // 0.008% today. The tolerance is the unit's own rounding, not a fudge: an
    // inset is a whole number of thousandths of a percent.
    expect(worstCover, "a cover crop is not showing the shape's ratio").toBeLessThan(0.0002);
    expect(worstContain, "a contained image is not keeping its own ratio").toBeLessThan(0.0002);
  });

  it("runs out of unit before it runs out of sense", () => {
    /**
     * The boundary, pinned rather than fixed. An inset is a whole number of
     * thousandths of a percent, so once a cover crop keeps less than 1/100000
     * of the image the two sides sum to the whole width and the source rect is
     * empty — the picture would not draw.
     *
     * It needs a ratio mismatch of about a hundred thousand: a one-unit-wide
     * shape holding a 1920x1 image. A 1000x1 spacer in an ordinary 200x100 box
     * is a mismatch of 500 and still keeps a visible sliver, so nothing a deck
     * plausibly holds reaches this.
     *
     * Written down because the arithmetic above is otherwise exact, and the one
     * place it stops being exact should not be a surprise to the next reader.
     */
    const empty = coverSrcRect({ w: 1, h: 1000 }, { w: 1920, h: 1 });
    expect(empty.l + empty.r).toBe(M);

    // And the case one might mistake for it, which is fine.
    const sliver = coverSrcRect({ w: 200, h: 100 }, { w: 1000, h: 1 });
    expect(sliver.l + sliver.r).toBeLessThan(M);
  });
});

describe("a shape inside a group somebody has resized", () => {
  /**
   * A child of a `<p:grpSp>` states its size in the group's CHILD coordinate
   * space, and the group scales that space by `ext ÷ chExt`. PowerPoint writes
   * the two equal for a new group and leaves `chExt` alone when the user drags
   * the group's handles, so every group anybody has resized has a scale factor.
   *
   * `shapeBox` read the declared extent and called it the rendered box, so
   * `cover` — whose whole job is not to distort — computed its crop for the
   * wrong ratio. Nothing reported it: a crop was computed and written.
   */
  /** The innermost `<p:sp>` of a spTree built from this markup. */
  function shapeIn(markup: string): Element {
    const doc = parseXml(
      `<p:sld xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:cSld><p:spTree>${markup}</p:spTree></p:cSld></p:sld>`,
    );
    return elements(doc, P_NS, "sp").at(-1) ?? doc.documentElement;
  }

  const inner = `<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:sp>`;
  const group = (cx: number, cy: number, chx: number, chy: number, body: string) =>
    `<p:grpSp><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="${chx}" cy="${chy}"/></a:xfrm></p:grpSpPr>${body}</p:grpSp>`;

  it("answers the box the picture is actually seen in", () => {
    // Stretched 4x across, so a 100x100 child is rendered 400x100.
    expect(shapeBox(shapeIn(group(400, 100, 100, 100, inner)))).toEqual({ w: 400, h: 100 });
  });

  it("crops the axis the rendered box asks for, not the declared one", () => {
    const box = shapeBox(shapeIn(group(400, 100, 100, 100, inner)));
    // A 2:1 picture in a box that is rendered 4:1: the picture is proportionally
    // TALLER than its box, so top and bottom go. Read off the declared 1:1 box
    // it was the sides, which then got stretched 4:1 by the group — a squashed
    // photo on the one run that asked not to be squashed.
    expect(coverSrcRect(box!, { w: 64, h: 32 })).toEqual({ l: 0, t: 25000, r: 0, b: 25000 });
  });

  it("accumulates through nested groups", () => {
    const nested = group(400, 100, 100, 100, group(100, 200, 100, 100, inner));
    expect(shapeBox(shapeIn(nested))).toEqual({ w: 400, h: 200 });
  });

  it("leaves the box alone for a group that has not been resized", () => {
    // What PowerPoint writes for a new group, and the case that must not move.
    expect(shapeBox(shapeIn(group(100, 100, 100, 100, inner)))).toEqual({ w: 100, h: 100 });
  });

  it("keeps the declared box when a group's own numbers cannot be read", () => {
    const noChExt = `<p:grpSp><p:grpSpPr><a:xfrm><a:ext cx="400" cy="100"/></a:xfrm></p:grpSpPr>${inner}</p:grpSp>`;
    expect(shapeBox(shapeIn(noChExt))).toEqual({ w: 100, h: 100 });
  });
});
