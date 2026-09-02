/**
 * The rotation a photo asks for, and what to do about it.
 *
 * A phone held upright writes the pixels LANDSCAPE and sets an EXIF tag saying
 * "turn this 90 degrees to show it". Every phone gallery and every browser
 * honours that tag.
 *
 * **PowerPoint does not.** That was an open question in `docs/TEST-KIT.md` for
 * weeks, because the two possible answers wanted opposite fixes and only a real
 * host could say which. The round of 2026-09-02 answered it on PowerPoint for
 * the web: a portrait photo with `Orientation=6` merged into a portrait frame
 * came out **lying on its side**, with the top and bottom of the subject
 * cropped off the left and right edges. So PowerPoint reads the stored pixels
 * and ignores the tag.
 *
 * That settles the fix. Swapping the dimensions this add-in reports would make
 * it worse — the crop would be computed for a picture the host is not drawing.
 * The bytes themselves have to be turned before they go into the package, and
 * then the tag is a lie about an image that no longer needs it.
 *
 * This file decides; it does not rotate. Turning pixels needs a decoder, and
 * `core` runs in the suite where there is no DOM — the same reason `read.ts`
 * parses headers by hand rather than asking `createImageBitmap`. The pane does
 * the turning, at the one place it reads a picked file.
 */

/**
 * An EXIF orientation, 1 to 8.
 *
 * 1 is upright. 3 is upside down. 6 and 8 are the quarter turns a phone writes
 * when it is held in portrait, and they are the two that matter here: they are
 * the ones where the stored pixels and the intended picture have their width
 * and height the other way round.
 *
 * 2, 4, 5 and 7 are the mirrored ones. They exist, they are rare outside
 * scanners, and they are reported rather than quietly treated as upright.
 */
export type Orientation = 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8;

/** What has to happen to the stored pixels to make the picture upright. */
export interface Correction {
  /** Clockwise degrees to turn the stored pixels. */
  rotate: 0 | 90 | 180 | 270;
  /** Whether the picture is mirrored as well as turned. */
  flip: boolean;
  /**
   * Whether width and height swap once it is turned.
   *
   * The crop maths in `place.ts` reads the dimensions out of the file header,
   * so this is the difference between a frame filled correctly and one filled
   * with a picture rotated a quarter turn.
   */
  swapsAxes: boolean;
}

const CORRECTIONS: Record<Orientation, Correction> = {
  1: { rotate: 0, flip: false, swapsAxes: false },
  2: { rotate: 0, flip: true, swapsAxes: false },
  3: { rotate: 180, flip: false, swapsAxes: false },
  4: { rotate: 180, flip: true, swapsAxes: false },
  5: { rotate: 90, flip: true, swapsAxes: true },
  6: { rotate: 90, flip: false, swapsAxes: true },
  7: { rotate: 270, flip: true, swapsAxes: true },
  8: { rotate: 270, flip: false, swapsAxes: true },
};

/** What to do about an orientation. Upright and unknown both mean "nothing". */
export function correctionFor(orientation: number | undefined): Correction {
  // `Number.isInteger`, not just a range. 1.5 satisfies `>= 1 && <= 8` and then
  // indexes nothing, so this returned undefined from a function whose type says
  // it cannot — and the caller read `.rotate` off it and threw. The same trap
  // `sweepPlan` documents: two of three quantities checked and the third
  // trusted. Found by a test that passed 1.5 among the nonsense values.
  const known = orientation !== undefined && Number.isInteger(orientation) && orientation >= 1 && orientation <= 8;
  return known ? CORRECTIONS[orientation as Orientation] : CORRECTIONS[1];
}

/** Whether this picture needs turning before a host that ignores the tag. */
export function needsTurning(orientation: number | undefined): boolean {
  const c = correctionFor(orientation);
  return c.rotate !== 0 || c.flip;
}

/**
 * The EXIF orientation in a JPEG's APP1 segment, or undefined.
 *
 * Undefined for every image that does not carry one, which is most of them:
 * PNG, GIF and BMP have no EXIF at all, and a JPEG straight out of an editor
 * usually has the tag stripped. Undefined is also the answer for a malformed
 * segment — this is a photograph, not a contract, and a merge that met a
 * strange file must still place it rather than stop.
 */
export function jpegOrientation(bytes: Uint8Array): Orientation | undefined {
  if (bytes[0] !== 0xff || bytes[1] !== 0xd8) return undefined; // not a JPEG
  let at = 2;
  while (at + 3 < bytes.length) {
    if (bytes[at] !== 0xff) {
      at++;
      continue;
    }
    const marker = bytes[at + 1] ?? 0;
    // Padding, and the standalone markers that carry no length. Skipping these
    // by a length reads two bytes of image data as one; `read.ts` carries the
    // same list for the same reason.
    if (marker === 0xff) {
      at++;
      continue;
    }
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      at += 2;
      continue;
    }
    // Entropy-coded data starts here and EXIF cannot appear after it.
    if (marker === 0xd9 || marker === 0xda) return undefined;

    const length = ((bytes[at + 2] ?? 0) << 8) | (bytes[at + 3] ?? 0);
    if (length < 2) return undefined;
    if (marker === 0xe1) {
      const found = inExifSegment(bytes, at + 4, at + 2 + length);
      if (found !== undefined) return found;
    }
    at += 2 + length;
  }
  return undefined;
}

/** `Exif\0\0`, then a TIFF header, then IFD0. */
function inExifSegment(bytes: Uint8Array, from: number, end: number): Orientation | undefined {
  const EXIF = [0x45, 0x78, 0x69, 0x66, 0x00, 0x00];
  if (EXIF.some((b, i) => bytes[from + i] !== b)) return undefined;

  const tiff = from + 6;
  if (tiff + 8 > end) return undefined;
  // "II" is little-endian, "MM" big. Both occur in the wild: phones are mostly
  // little, some cameras big, and reading one as the other turns tag 0x0112
  // into 0x1201 and finds nothing.
  const little = bytes[tiff] === 0x49 && bytes[tiff + 1] === 0x49;
  const big = bytes[tiff] === 0x4d && bytes[tiff + 1] === 0x4d;
  if (!little && !big) return undefined;
  const u16 = (at: number) =>
    little ? (bytes[at] ?? 0) | ((bytes[at + 1] ?? 0) << 8) : ((bytes[at] ?? 0) << 8) | (bytes[at + 1] ?? 0);
  const u32 = (at: number) =>
    little
      ? ((bytes[at] ?? 0) |
          ((bytes[at + 1] ?? 0) << 8) |
          ((bytes[at + 2] ?? 0) << 16) |
          ((bytes[at + 3] ?? 0) << 24)) >>>
        0
      : (((bytes[at] ?? 0) << 24) |
          ((bytes[at + 1] ?? 0) << 16) |
          ((bytes[at + 2] ?? 0) << 8) |
          (bytes[at + 3] ?? 0)) >>>
        0;

  if (u16(tiff + 2) !== 0x002a) return undefined;
  const ifd0 = tiff + u32(tiff + 4);
  if (ifd0 + 2 > end) return undefined;

  const entries = u16(ifd0);
  for (let i = 0; i < entries; i++) {
    const entry = ifd0 + 2 + i * 12;
    if (entry + 12 > end) return undefined;
    if (u16(entry) !== 0x0112) continue;
    // A SHORT sits in the first two bytes of the four-byte value field, and
    // those two are at the START of it in both byte orders.
    const value = u16(entry + 8);
    return value >= 1 && value <= 8 ? (value as Orientation) : undefined;
  }
  return undefined;
}
