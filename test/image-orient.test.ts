import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { correctionFor, jpegOrientation, needsTurning } from "../src/core/image/orient.js";
import { readImage } from "../src/core/image/read.js";

/**
 * The EXIF orientation reader.
 *
 * Checked against REAL files, written by Pillow the way a camera writes them —
 * `scripts/build-exif-fixtures.py` makes them. A decoder tested only against
 * bytes the test assembled proves that the encoder and the decoder agree with
 * each other, which is the one thing nobody needed to know.
 */
const FIX = "test/fixtures/exif";
const read = (name: string) => new Uint8Array(readFileSync(`${FIX}/${name}`));

describe("reading a photo's EXIF orientation", () => {
  it.each([1, 2, 3, 4, 5, 6, 7, 8])("finds orientation %i in a real JPEG", (tag) => {
    expect(jpegOrientation(read(`orientation-${tag}.jpg`))).toBe(tag);
  });

  it("answers nothing for a JPEG with no EXIF", () => {
    expect(jpegOrientation(read("no-exif.jpg"))).toBeUndefined();
  });

  it("answers nothing for a PNG, which cannot carry EXIF", () => {
    expect(jpegOrientation(read("plain.png"))).toBeUndefined();
  });

  it("answers nothing rather than throwing on a truncated file", () => {
    // A merge meets whatever the user's folder holds. One unreadable picture
    // must leave one placeholder visible, not end the run.
    const whole = read("orientation-6.jpg");
    for (const cut of [2, 8, 20, 40, whole.length - 1]) {
      expect(() => jpegOrientation(whole.subarray(0, cut))).not.toThrow();
    }
  });

  it("reads a BIG-endian TIFF header as well as a little-endian one", () => {
    /**
     * Pillow writes little-endian, so the fixtures above only exercise "II".
     * Some cameras write "MM", and reading one as the other turns tag 0x0112
     * into 0x1201 and finds nothing — a silent wrong answer, not a crash.
     *
     * Built by hand because the generator cannot produce it, and kept minimal:
     * SOI, APP1 with one IFD0 entry, EOI.
     */
    const be = [
      0xff,
      0xd8, // SOI
      0xff,
      0xe1,
      0x00,
      0x20, // APP1, length 32
      0x45,
      0x78,
      0x69,
      0x66,
      0x00,
      0x00, // "Exif\0\0"
      0x4d,
      0x4d,
      0x00,
      0x2a,
      0x00,
      0x00,
      0x00,
      0x08, // MM, 42, IFD0 at 8
      0x00,
      0x01, // one entry
      0x01,
      0x12,
      0x00,
      0x03,
      0x00,
      0x00,
      0x00,
      0x01,
      0x00,
      0x08,
      0x00,
      0x00, // tag 0x0112 = 8
      0x00,
      0x00,
      0x00,
      0x00, // no next IFD
      0xff,
      0xd9, // EOI
    ];
    expect(jpegOrientation(new Uint8Array(be))).toBe(8);
  });

  it("does not disturb what readImage already answers", () => {
    // The orientation reader walks the same segment chain. If it had been put
    // inside `readImage` and got the walk wrong, the dimensions would go with
    // it; they are separate functions and this says so.
    const info = readImage(read("orientation-6.jpg"));
    expect(info?.kind).toBe("jpeg");
    expect(info?.width).toBe(8);
    expect(info?.height).toBe(4);
  });
});

describe("deciding what to do about an orientation", () => {
  it("does nothing for upright, missing, or nonsense", () => {
    for (const value of [1, undefined, 0, 9, -1, 1.5, NaN]) {
      expect(needsTurning(value), `orientation ${value}`).toBe(false);
      expect(correctionFor(value).rotate).toBe(0);
    }
  });

  it("turns the two quarter-turns a phone actually writes", () => {
    expect(correctionFor(6)).toEqual({ rotate: 90, flip: false, swapsAxes: true });
    expect(correctionFor(8)).toEqual({ rotate: 270, flip: false, swapsAxes: true });
  });

  it("knows which orientations swap width and height", () => {
    // The ones that swap are exactly the quarter turns. `place.ts` computes the
    // crop from the header's dimensions, so getting this wrong fills the frame
    // with a picture a quarter turn out — which is the defect the real-host
    // round of 2026-09-02 photographed.
    const swaps = [1, 2, 3, 4, 5, 6, 7, 8].filter((o) => correctionFor(o).swapsAxes);
    expect(swaps).toEqual([5, 6, 7, 8]);
  });

  it("calls every orientation except upright a turn", () => {
    const turning = [1, 2, 3, 4, 5, 6, 7, 8].filter((o) => needsTurning(o));
    expect(turning, "2 and 4 are mirrored without rotation, and still need work").toEqual([2, 3, 4, 5, 6, 7, 8]);
  });
});
