/**
 * What an image IS, read from its own first bytes.
 *
 * The merge needs two things from a picture before it can place one: the
 * content type, so the package can declare the part, and the pixel dimensions,
 * so the fill can cover or contain the shape without distorting. Both are in
 * the file's header, and neither can be taken from the FILE NAME — a `.jpg`
 * that is really a PNG is an ordinary thing to find in a folder of exports, and
 * declaring the wrong content type is how a deck opens as damaged.
 *
 * Pure and byte-level on purpose. `Image`/`createImageBitmap` would answer the
 * same question, and would tie the engine to a browser: the same merge has to
 * run in the suite, where there is no DOM.
 */

/** The formats a merge will embed. PowerPoint renders all four everywhere. */
export type ImageKind = "png" | "jpeg" | "gif" | "bmp";

export interface ImageInfo {
  kind: ImageKind;
  /** The OOXML content type for the part. */
  contentType: string;
  /** The extension a media part is named with, matching `contentType`. */
  extension: string;
  width: number;
  height: number;
}

const CONTENT_TYPES: Record<ImageKind, string> = {
  png: "image/png",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
};

const EXTENSIONS: Record<ImageKind, string> = { png: "png", jpeg: "jpeg", gif: "gif", bmp: "bmp" };

function u16be(b: Uint8Array, at: number): number {
  return ((b[at] ?? 0) << 8) | (b[at + 1] ?? 0);
}
function u32be(b: Uint8Array, at: number): number {
  return (((b[at] ?? 0) << 24) | ((b[at + 1] ?? 0) << 16) | ((b[at + 2] ?? 0) << 8) | (b[at + 3] ?? 0)) >>> 0;
}
function u16le(b: Uint8Array, at: number): number {
  return (b[at] ?? 0) | ((b[at + 1] ?? 0) << 8);
}
function i32le(b: Uint8Array, at: number): number {
  return (b[at] ?? 0) | ((b[at + 1] ?? 0) << 8) | ((b[at + 2] ?? 0) << 16) | ((b[at + 3] ?? 0) << 24) | 0;
}

function starts(bytes: Uint8Array, signature: number[]): boolean {
  return signature.every((byte, i) => bytes[i] === byte);
}

/**
 * A JPEG's dimensions, which are not at a fixed offset.
 *
 * A JPEG is a chain of segments and the size lives in whichever start-of-frame
 * marker the encoder used — SOF0 for baseline, SOF2 for progressive, and a
 * dozen others. So the chain is walked rather than indexed into. The markers
 * that are NOT frame headers have to be skipped by their own length, and four
 * of them (`D0`-`D9`, `01`) carry no length at all: treating those as
 * length-bearing walks off into the entropy-coded data and reads two arbitrary
 * bytes as a width.
 */
function jpegSize(bytes: Uint8Array): { width: number; height: number } | undefined {
  let at = 2; // past SOI
  while (at + 3 < bytes.length) {
    if (bytes[at] !== 0xff) {
      at++; // fill byte, or padding between segments
      continue;
    }
    const marker = bytes[at + 1] ?? 0;
    // Any number of 0xFF bytes may pad the space before a marker, and the
    // standard says so. Taken as a marker itself, `FF FF C0 ...` reads the
    // FRAME HEADER's own first bytes as a segment length, skips a nonsense
    // distance, runs off the end and reports a perfectly good JPEG as
    // unreadable — one placeholder left visible with nothing to say why.
    if (marker === 0xff) {
      at++;
      continue;
    }
    // Standalone markers: no length follows, so nothing may be skipped by one.
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      at += 2;
      continue;
    }
    if (marker === 0xd9 || marker === 0xda) return undefined; // end, or entropy data
    const length = u16be(bytes, at + 2);
    if (length < 2) return undefined;
    // Every SOFn except the four that are not frame headers: DHT (C4),
    // JPG (C8) and DAC (CC) sit inside the same C0-CF range.
    const isFrame = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isFrame) {
      // SOF payload: precision(1), height(2), width(2).
      if (at + 9 >= bytes.length) return undefined;
      return { height: u16be(bytes, at + 5), width: u16be(bytes, at + 7) };
    }
    at += 2 + length;
  }
  return undefined;
}

/**
 * Read an image's kind and size, or answer undefined.
 *
 * Undefined for anything it does not recognise, and for a file whose header is
 * truncated or nonsense — never a guess and never a throw. A merge meets
 * whatever the user's folder holds, and one unreadable file must leave one
 * placeholder visible rather than end the run.
 */
export function readImage(bytes: Uint8Array): ImageInfo | undefined {
  const of = (kind: ImageKind, width: number, height: number): ImageInfo | undefined =>
    width > 0 && height > 0
      ? { kind, contentType: CONTENT_TYPES[kind], extension: EXTENSIONS[kind], width, height }
      : undefined;

  if (starts(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) {
    // IHDR is required to be the first chunk, so width and height are at a
    // fixed offset. Checked rather than assumed: a PNG whose first chunk is
    // something else is not a PNG this can read.
    if (bytes.length < 24) return undefined;
    if (String.fromCharCode(...bytes.slice(12, 16)) !== "IHDR") return undefined;
    return of("png", u32be(bytes, 16), u32be(bytes, 20));
  }
  if (starts(bytes, [0xff, 0xd8, 0xff])) {
    const size = jpegSize(bytes);
    return size ? of("jpeg", size.width, size.height) : undefined;
  }
  if (starts(bytes, [0x47, 0x49, 0x46, 0x38])) {
    if (bytes.length < 10) return undefined;
    return of("gif", u16le(bytes, 6), u16le(bytes, 8));
  }
  if (starts(bytes, [0x42, 0x4d])) {
    if (bytes.length < 22) return undefined;
    // WHICH header this is decides where the size lives, and the header states
    // its own length at offset 14. The 12-byte OS/2 header keeps width and
    // height as two 16-bit numbers at 18 and 20; every later one keeps them as
    // two 32-bit numbers at 18 and 22.
    //
    // Reading the second shape out of the first does not fail, which is what
    // makes it worth a branch: a 200 x 100 bitmap came back as 6553800 x
    // 1572865, because `200 | (100 << 16)` is a number. The picture is then
    // cropped to a ratio that has nothing to do with it, and nothing anywhere
    // says the size was invented.
    const header = i32le(bytes, 14);
    if (header === 12) return of("bmp", u16le(bytes, 18), u16le(bytes, 20));
    // A header shape this does not know is refused rather than guessed at.
    if (header < 16) return undefined;
    if (bytes.length < 26) return undefined;
    // BMP stores height NEGATIVE for a top-down bitmap, which is a direction
    // rather than a size. Taken raw it is a negative extent, and every ratio
    // built from it comes out backwards.
    return of("bmp", Math.abs(i32le(bytes, 18)), Math.abs(i32le(bytes, 22)));
  }
  return undefined;
}
