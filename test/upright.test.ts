import { readFileSync } from "node:fs";
import { afterEach, describe, expect, it, vi } from "vitest";
import { upright } from "../src/pane/upright.js";

/**
 * Turning a photo's pixels before the deck sees them.
 *
 * PowerPoint ignores EXIF orientation, so a phone's portrait photo arrives on
 * its side unless the bytes are turned. jsdom has neither `createImageBitmap`
 * nor a canvas, which is useful in itself: it is exactly the environment the
 * fail-soft path exists for, and the first four cases below run in it with no
 * stubbing at all.
 */
const read = (name: string) => new Uint8Array(readFileSync(`test/fixtures/exif/${name}`));

const REAL = {
  createImageBitmap: (globalThis as Record<string, unknown>).createImageBitmap,
  OffscreenCanvas: (globalThis as Record<string, unknown>).OffscreenCanvas,
};

afterEach(() => {
  const g = globalThis as Record<string, unknown>;
  g.createImageBitmap = REAL.createImageBitmap;
  g.OffscreenCanvas = REAL.OffscreenCanvas;
  vi.restoreAllMocks();
});

describe("leaving a picture alone", () => {
  it("returns a PNG untouched, because it cannot carry EXIF", async () => {
    const bytes = read("plain.png");
    expect(await upright(bytes)).toBe(bytes);
  });

  it("returns a JPEG with no EXIF untouched", async () => {
    const bytes = read("no-exif.jpg");
    expect(await upright(bytes)).toBe(bytes);
  });

  it("returns an upright JPEG untouched, without asking for a canvas", async () => {
    // The common case. It must not cost a decode.
    const g = globalThis as Record<string, unknown>;
    const bitmap = vi.fn();
    g.createImageBitmap = bitmap;
    const bytes = read("orientation-1.jpg");
    expect(await upright(bytes)).toBe(bytes);
    expect(bitmap, "an upright photo needs no decoding").not.toHaveBeenCalled();
  });

  it("returns the photo unchanged where there is no way to draw", async () => {
    // A merge that stops because a canvas was unavailable is worse than a photo
    // the wrong way up. This is jsdom, so nothing is stubbed: the guard is
    // being exercised by the environment itself.
    const bytes = read("orientation-6.jpg");
    expect(await upright(bytes)).toBe(bytes);
  });
});

describe("turning a photo that asks for it", () => {
  /** A canvas that records what was done to it rather than drawing. */
  function stubCanvas(bitmapSize: { width: number; height: number }) {
    const calls: string[] = [];
    const ctx = {
      translate: (x: number, y: number) => calls.push(`translate ${x},${y}`),
      rotate: (r: number) => calls.push(`rotate ${Math.round((r * 180) / Math.PI)}`),
      scale: (x: number, y: number) => calls.push(`scale ${x},${y}`),
      drawImage: () => calls.push("drawImage"),
    };
    const made: { width: number; height: number }[] = [];
    const g = globalThis as Record<string, unknown>;
    g.createImageBitmap = vi.fn(() => Promise.resolve({ ...bitmapSize, close: () => {} }));
    g.OffscreenCanvas = class {
      constructor(
        public width: number,
        public height: number,
      ) {
        made.push({ width, height });
      }
      getContext() {
        return ctx;
      }
      convertToBlob() {
        return Promise.resolve({ arrayBuffer: () => Promise.resolve(new Uint8Array([1, 2, 3]).buffer) });
      }
    };
    return { calls, made };
  }

  it("turns orientation 6 a quarter turn clockwise and swaps the axes", async () => {
    // The phone-in-portrait case, and the one the real-host round photographed:
    // 1200x800 stored, meant to be seen as 800x1200.
    const { calls, made } = stubCanvas({ width: 1200, height: 800 });
    const out = await upright(read("orientation-6.jpg"));

    expect(made[0], "the canvas is the picture the right way up").toEqual({ width: 800, height: 1200 });
    expect(calls).toContain("rotate 90");
    expect(calls, "6 is a turn, not a mirror").not.toContain("scale -1,1");
    expect(Array.from(out), "the re-encoded bytes, not the originals").toEqual([1, 2, 3]);
  });

  it("turns orientation 8 the other way", async () => {
    const { calls, made } = stubCanvas({ width: 1200, height: 800 });
    await upright(read("orientation-8.jpg"));
    expect(made[0]).toEqual({ width: 800, height: 1200 });
    expect(calls).toContain("rotate 270");
  });

  it("mirrors the mirrored ones, and keeps their axes", async () => {
    // 2 is a mirror with no rotation. Nobody checks these by hand, and an extra
    // flip is invisible until somebody's scan comes out backwards.
    const { calls, made } = stubCanvas({ width: 1200, height: 800 });
    await upright(read("orientation-2.jpg"));
    expect(made[0], "no rotation, so no swap").toEqual({ width: 1200, height: 800 });
    expect(calls).toContain("rotate 0");
    expect(calls).toContain("scale -1,1");
  });

  it("keeps the original bytes when the canvas gives nothing back", async () => {
    const g = globalThis as Record<string, unknown>;
    g.createImageBitmap = vi.fn(() => Promise.resolve({ width: 1200, height: 800, close: () => {} }));
    g.OffscreenCanvas = class {
      getContext() {
        return null; // a host that refuses a 2d context
      }
    };
    const bytes = read("orientation-6.jpg");
    expect(await upright(bytes)).toBe(bytes);
  });

  it("keeps the original bytes when decoding throws", async () => {
    const g = globalThis as Record<string, unknown>;
    g.createImageBitmap = vi.fn(() => Promise.reject(new Error("not an image this browser will decode")));
    g.OffscreenCanvas = class {};
    const bytes = read("orientation-6.jpg");
    expect(await upright(bytes)).toBe(bytes);
  });
});

describe("the pane actually uses it", () => {
  /**
   * A source guard, and it says so rather than pretending to be a behaviour
   * test. jsdom has no canvas, so `upright` returns the bytes it was given and
   * a picked photo looks identical whether the pane routes through it or not —
   * there is nothing to observe. What can go wrong is somebody simplifying that
   * line back to a bare `arrayBuffer()`, and this is what notices.
   */
  it("passes every picked file through upright", () => {
    const main = readFileSync("src/pane/main.ts", "utf8");
    const line = main.split(/\r?\n/).find((l) => l.includes("images.set(file.name"));
    expect(line, "the picker no longer stores files this way; move this guard").toBeTruthy();
    expect(line, "picked bytes must be turned before they are held").toContain("upright(");
  });
});
