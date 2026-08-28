/**
 * How a picture sits in a shape, as numbers.
 *
 * Separated from the XML that carries it because this is the whole of the
 * decision and it is arithmetic: given a shape's box and an image's pixels,
 * which part of the image shows, and where. Kept pure so every case can be
 * checked without a package — including the ones nobody would think to open a
 * deck for, like an image one pixel wide.
 *
 * The unit is OOXML's own: `srcRect` and `fillRect` insets are THOUSANDTHS OF A
 * PERCENT, so a tenth of the image is 10000. They are also INSETS, not
 * coordinates — `l="10000"` means "start a tenth of the way in", and all four
 * are measured inward from their own edge. Getting that backwards mirrors the
 * picture rather than failing, which is the kind of wrong that ships.
 */

/** Where the image is anchored when cover crops it, or contain leaves space. */
export type FillMode = "cover" | "contain" | "stretch";

export interface Inset {
  l: number;
  t: number;
  r: number;
  b: number;
}

/** Nothing trimmed and nothing inset: the image fills the box exactly. */
export const NO_INSET: Inset = { l: 0, t: 0, r: 0, b: 0 };

const PER_CENT_MILLE = 100000;

/**
 * Round an inset pair so the two sides always sum to the same total.
 *
 * Rounding each side on its own drifts: 33333 and 33334 leave a pixel of the
 * shape uncovered, which on a photo against a coloured background is a hairline
 * of the wrong colour down one edge. Taking the FIRST side and deriving the
 * second from the total keeps the pair exact.
 */
function split(total: number): [number, number] {
  const first = Math.round(total / 2);
  return [first, total - first];
}

/**
 * The `srcRect` for COVER: which part of the image to show.
 *
 * The image is scaled until it covers the box, and the overflow is trimmed off
 * the long axis, centred. The trim is expressed as a share of the IMAGE, which
 * is why it is computed from the ratio of ratios rather than from pixels: what
 * is cropped is "how much wider the image is than the box", and that is a
 * proportion whatever the sizes.
 */
export function coverSrcRect(shape: { w: number; h: number }, image: { w: number; h: number }): Inset {
  if (shape.w <= 0 || shape.h <= 0 || image.w <= 0 || image.h <= 0) return NO_INSET;
  const shapeRatio = shape.w / shape.h;
  const imageRatio = image.w / image.h;
  if (Math.abs(shapeRatio - imageRatio) < 1e-9) return NO_INSET;
  if (imageRatio > shapeRatio) {
    // Image is proportionally wider: trim the sides.
    const keep = shapeRatio / imageRatio;
    const [l, r] = split(Math.round((1 - keep) * PER_CENT_MILLE));
    return { l, t: 0, r, b: 0 };
  }
  // Image is proportionally taller: trim top and bottom.
  const keep = imageRatio / shapeRatio;
  const [t, b] = split(Math.round((1 - keep) * PER_CENT_MILLE));
  return { l: 0, t, r: 0, b };
}

/**
 * The `fillRect` for CONTAIN: where inside the box the whole image sits.
 *
 * The mirror of cover, and it insets rather than crops — the image is scaled
 * until it fits and centred in what is left over. The space around it is
 * whatever the shape itself shows, which is why a shape with a fill of its own
 * letterboxes in that colour and a shape with none letterboxes transparently.
 * That is the shape's business and deliberately not decided here.
 */
export function containFillRect(shape: { w: number; h: number }, image: { w: number; h: number }): Inset {
  if (shape.w <= 0 || shape.h <= 0 || image.w <= 0 || image.h <= 0) return NO_INSET;
  const shapeRatio = shape.w / shape.h;
  const imageRatio = image.w / image.h;
  if (Math.abs(shapeRatio - imageRatio) < 1e-9) return NO_INSET;
  if (imageRatio > shapeRatio) {
    // Image is proportionally wider: it spans the width, bars above and below.
    const used = shapeRatio / imageRatio;
    const [t, b] = split(Math.round((1 - used) * PER_CENT_MILLE));
    return { l: 0, t, r: 0, b };
  }
  const used = imageRatio / shapeRatio;
  const [l, r] = split(Math.round((1 - used) * PER_CENT_MILLE));
  return { l, t: 0, r, b: 0 };
}
