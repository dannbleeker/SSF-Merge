/**
 * Turn a photo's pixels the way its EXIF tag asks, before it goes in the deck.
 *
 * PowerPoint ignores EXIF orientation. That was an open question for weeks —
 * the two possible answers wanted opposite fixes — and the real-host round of
 * 2026-09-02 settled it: a phone's portrait photo, `Orientation=6`, merged into
 * a portrait frame came out lying on its side, with the subject's head and feet
 * cropped off the left and right edges.
 *
 * So the bytes have to be turned here, and the tag is then a lie about a
 * picture that no longer needs it. `core/image/orient.ts` decides what turn is
 * wanted; this does it, and it lives in the pane rather than in `core` for the
 * same reason `read.ts` parses headers by hand: turning pixels needs a decoder,
 * and the engine runs in a suite with no DOM.
 *
 * Everything here fails SOFT. A picture that cannot be decoded, in a browser
 * without `createImageBitmap`, or under a canvas the host refuses, comes back
 * exactly as it arrived. A photo the wrong way up is a bad merge; a merge that
 * stops because a canvas was unavailable is a worse one.
 */
import { correctionFor, jpegOrientation, needsTurning } from "../core/image/orient.js";

/** Whether this environment can turn pixels at all. */
function canDraw(): boolean {
  return (
    typeof createImageBitmap === "function" &&
    (typeof OffscreenCanvas === "function" || typeof document !== "undefined")
  );
}

function surface(width: number, height: number): { canvas: OffscreenCanvas | HTMLCanvasElement; ctx: unknown } | null {
  if (typeof OffscreenCanvas === "function") {
    const canvas = new OffscreenCanvas(width, height);
    return { canvas, ctx: canvas.getContext("2d") };
  }
  if (typeof document === "undefined") return null;
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  return { canvas, ctx: canvas.getContext("2d") };
}

async function toBytes(canvas: OffscreenCanvas | HTMLCanvasElement, type: string): Promise<Uint8Array | null> {
  if ("convertToBlob" in canvas) {
    const blob = await canvas.convertToBlob({ type });
    return new Uint8Array(await blob.arrayBuffer());
  }
  const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, type));
  return blob ? new Uint8Array(await blob.arrayBuffer()) : null;
}

/**
 * The picture as it should be seen, or the bytes unchanged.
 *
 * Only JPEGs carry EXIF, and only some of those carry a turn, so the common
 * case costs one header read and returns the original array untouched.
 */
export async function upright(bytes: Uint8Array): Promise<Uint8Array> {
  const orientation = jpegOrientation(bytes);
  if (!needsTurning(orientation)) return bytes;
  if (!canDraw()) return bytes;

  const { rotate, flip, swapsAxes } = correctionFor(orientation);
  try {
    // `.slice()` rather than the array itself: a `Uint8Array` may sit over a
    // `SharedArrayBuffer`, which is not a `BlobPart`, and the copy is the one
    // way to say "plain ArrayBuffer" that both tsc and the linter accept. It is
    // one copy per picked file, against a decode and a re-encode either side.
    const bitmap = await createImageBitmap(new Blob([bytes.slice()], { type: "image/jpeg" }));
    const width = swapsAxes ? bitmap.height : bitmap.width;
    const height = swapsAxes ? bitmap.width : bitmap.height;

    const made = surface(width, height);
    const ctx = made?.ctx as CanvasRenderingContext2D | null | undefined;
    if (!made || !ctx) {
      bitmap.close?.();
      return bytes;
    }

    // Move to the middle, turn, mirror if the tag asked, then draw centred.
    // Done as one transform rather than a chain of ifs because the mirrored
    // orientations (2, 4, 5, 7) are the ones nobody tests by hand and the ones
    // where an extra flip is invisible until somebody's scan comes out
    // backwards.
    ctx.translate(width / 2, height / 2);
    ctx.rotate((rotate * Math.PI) / 180);
    if (flip) ctx.scale(-1, 1);
    ctx.drawImage(bitmap, -bitmap.width / 2, -bitmap.height / 2);
    bitmap.close?.();

    const turned = await toBytes(made.canvas, "image/jpeg");
    // Re-encoded without the EXIF block, so the tag cannot ask for the turn a
    // second time in a viewer that DOES honour it — which would leave the
    // picture wrong in exactly the places it used to be right.
    return turned ?? bytes;
  } catch {
    return bytes;
  }
}
