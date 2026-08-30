#!/usr/bin/env node
/**
 * The ribbon icons, drawn in code rather than checked in as binaries.
 *
 * Office wants 16, 32, 64 and 80 pixel PNGs and a manifest that names each one
 * by URL. Four hand-made files is four things to keep in step with a palette
 * that lives in `src/pane/taskpane.css`, and a binary in a diff is a change
 * nobody can review. This writes them from the same two colours the pane uses,
 * so "the icon is off-brand" is a one-line change rather than a round trip
 * through an image editor.
 *
 * The encoder is a few dozen lines because a PNG of a flat shape does not need
 * a library: a raster, one zlib stream, three chunks and a CRC. `node:zlib`
 * does the only hard part.
 *
 *   node scripts/build-icons.mjs
 */
import { deflateSync } from "node:zlib";
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { isMain } from "./is-main.mjs";

/** The pane's own two colours. Navy ground, orange tick. */
export const NAVY = [0x00, 0x25, 0x4c];
export const ORANGE = [0xed, 0x89, 0x36];

const CRC = (() => {
  const table = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c;
  }
  return (buf) => {
    let c = -1;
    for (const b of buf) c = table[(c ^ b) & 0xff] ^ (c >>> 8);
    return (c ^ -1) >>> 0;
  };
})();

function chunk(type, body) {
  const head = Buffer.alloc(8);
  head.writeUInt32BE(body.length, 0);
  head.write(type, 4, "ascii");
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(CRC(Buffer.concat([head.subarray(4), body])), 0);
  return Buffer.concat([head, body, crc]);
}

/**
 * An RGBA raster as a PNG.
 *
 * `pixel(x, y)` answers `[r, g, b, a]`. Colour type 6 (RGBA) and filter 0 on
 * every row: a flat shape compresses to nothing either way, and a filter that
 * is always zero is one less thing to get wrong.
 */
export function png(size, pixel) {
  const raw = Buffer.alloc(size * (size * 4 + 1));
  let at = 0;
  for (let y = 0; y < size; y++) {
    raw[at++] = 0; // filter: none
    for (let x = 0; x < size; x++) {
      const [r, g, b, a] = pixel(x, y);
      raw[at++] = r;
      raw[at++] = g;
      raw[at++] = b;
      raw[at++] = a;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw, { level: 9 })),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

/**
 * The mark: a navy square holding one orange row and two pale copies of it.
 *
 * That IS the product — one template block, one set per row of your data — and
 * it is the only thing this icon has room to say at sixteen pixels. The first
 * version drew the pane's own tick, a single bar across the middle, and it read
 * as a minus sign; the second and third rows are what turn a dash into a
 * statement about repetition.
 *
 * Everything is in SIXTEENTHS of the icon rather than in pixels, so the same
 * arithmetic draws 16 and 80 and the proportions do not drift between them.
 */
export function markPixel(size) {
  const u = size / 16;
  const radius = Math.max(1, Math.round(size * 0.16));
  const left = Math.round(3.5 * u);
  const right = size - left;
  // Three rows on a four-unit pitch, each about two units tall — the tallest
  // that leaves a visible gap once 16px rounds them to whole pixels.
  const rows = [0, 1, 2].map((i) => {
    const top = Math.round((4 + i * 4) * u);
    return [top, Math.max(1, Math.round(top + 2 * u)) === top ? top + 1 : Math.round(top + 2 * u)];
  });

  const outside = (x, y) => {
    // Only the four corners can be outside a rounded square.
    const cx = x < radius ? radius - 0.5 : x > size - radius - 1 ? size - radius - 0.5 : x;
    const cy = y < radius ? radius - 0.5 : y > size - radius - 1 ? size - radius - 0.5 : y;
    if (cx === x && cy === y) return false;
    return Math.hypot(x - cx, y - cy) > radius;
  };

  return (x, y) => {
    if (outside(x, y)) return [0, 0, 0, 0];
    if (x < left || x >= right) return [...NAVY, 255];
    const row = rows.findIndex(([top, bottom]) => y >= top && y < bottom);
    if (row === 0) return [...ORANGE, 255];
    // The copies, in the pane's own pale blue. Not white: white rows on navy
    // read as a table, and the point is that they are the SAME row again.
    if (row > 0) return [0xdd, 0xeb, 0xf7, 255];
    return [...NAVY, 255];
  };
}

/** The sizes Office asks for. 64 is the unified manifest's colour icon. */
export const SIZES = [16, 32, 64, 80];

/**
 * The Marketplace icon, which is a listing upload rather than a manifest URL.
 *
 * Partner Center asks for 300x300 and at most 512 KB on the offer listing page,
 * and it is a separate thing from the two icons the manifest names. The Logo
 * section of `docs/listing/LISTING.md` said the logo needed no work, which was
 * true of the manifest pair and wrong about this.
 */
export const MARKETPLACE = 300;

/**
 * The mark, drawn at `factor` times the size and averaged back down.
 *
 * `markPixel`'s `outside` test is a hard boolean: a pixel is in the rounded
 * square or it is not. At sixteen pixels nobody can see the staircase on the
 * corners, and at three hundred, blown up on a listing page, it is the first
 * thing you see. Averaging 16 samples per pixel gives the corners the partial
 * coverage they should have had.
 *
 * The averaging is PREMULTIPLIED. Outside pixels are transparent black, so
 * averaging the colour channels raw drags every edge pixel toward black and
 * rings the icon in grey. Weighting each sample by its own alpha is the
 * difference between a smooth edge and a dirty one.
 *
 * The four sizes above are deliberately NOT drawn this way. They ship, they are
 * pinned byte-for-byte by `manifest.test.ts`, and a small icon that Office
 * scales again is better off with hard pixels than with soft ones.
 */
export function supersampled(size, factor = 4) {
  const fine = markPixel(size * factor);
  return (x, y) => {
    let r = 0;
    let g = 0;
    let b = 0;
    let a = 0;
    for (let dy = 0; dy < factor; dy++) {
      for (let dx = 0; dx < factor; dx++) {
        const [sr, sg, sb, sa] = fine(x * factor + dx, y * factor + dy);
        r += sr * sa;
        g += sg * sa;
        b += sb * sa;
        a += sa;
      }
    }
    if (a === 0) return [0, 0, 0, 0];
    return [Math.round(r / a), Math.round(g / a), Math.round(b / a), Math.round(a / (factor * factor))];
  };
}

export function main() {
  const root = dirname(dirname(fileURLToPath(import.meta.url)));
  const out = join(root, "public", "assets");
  mkdirSync(out, { recursive: true });
  for (const size of SIZES) {
    writeFileSync(join(out, `icon-${size}.png`), png(size, markPixel(size)));
  }
  // Beside the screenshots rather than in public/assets: it is uploaded to
  // Partner Center by hand and no manifest points at it, so serving it from the
  // site would put a file on the web that nothing ever requests.
  const listing = join(root, "docs", "listing");
  mkdirSync(listing, { recursive: true });
  writeFileSync(join(listing, `marketplace-icon-${MARKETPLACE}.png`), png(MARKETPLACE, supersampled(MARKETPLACE)));
  // The unified manifest wants a monochrome OUTLINE icon as well, and it is a
  // different picture rather than the same one recoloured: it is drawn on a
  // transparent ground and stencilled by the host, so a navy square would come
  // back as a solid block. The three rows are the whole mark that survives.
  writeFileSync(
    join(out, "icon-outline-32.png"),
    png(32, (x, y) => {
      const [r, g, b, a] = markPixel(32)(x, y);
      const isGround = r === NAVY[0] && g === NAVY[1] && b === NAVY[2];
      return a === 255 && !isGround ? [255, 255, 255, 255] : [0, 0, 0, 0];
    }),
  );
  console.log(
    `icons: ${SIZES.map((s) => `icon-${s}.png`).join(", ")}, icon-outline-32.png, ` +
      `marketplace-icon-${MARKETPLACE}.png`,
  );
}

if (isMain(import.meta.url)) main();
