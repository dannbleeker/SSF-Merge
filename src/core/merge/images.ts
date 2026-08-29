/**
 * The pass that turns `{{Photo|image}}` into an actual picture.
 *
 * Runs BEFORE the text pass and must: to `mergeParagraph` an image field is an
 * ordinary field with an unknown format, so left to itself it would write the
 * FILE NAME onto the slide as text. This pass takes the placeholder away and
 * fills the shape instead, and the text pass then finds nothing to do.
 *
 * Everything happens in the package. The alternative is `ShapeFill.setImage`,
 * which on PowerPoint for the web STRETCHES — measured 2026-08-28, a square
 * card into a 2:1 box comes out a wide ellipse — so going through the API would
 * mean letterboxing every picture by hand before sending it. Here the fill mode
 * is written directly and the host has no say.
 */
import { readImage } from "../image/read.js";
import { fillShapeWithImage, shapeOf, shapesIn } from "../image/place.js";
import type { FillMode } from "../image/fill.js";
import type { Pkg } from "../pptx/pkg.js";
import { A_NS, elements } from "../pptx/xml.js";
import { FIELD } from "./text.js";

const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/**
 * The formats that mean "this is a picture", and which fill they ask for.
 *
 * `image` covers, because a photo in a frame that letterboxes looks broken and
 * cropping is what every design tool does. `image-fit` contains, because
 * cropping a LOGO is wrong in a way cropping a photo is not — a trimmed
 * wordmark is somebody else's trademark, mangled.
 */
const MODES: Record<string, FillMode> = {
  image: "cover",
  "image-fit": "contain",
  "image-stretch": "stretch",
};

/** Whether a format spec asks for a picture rather than formatted text. */
export function imageMode(format: string | undefined): FillMode | undefined {
  if (format === undefined) return undefined;
  return Object.prototype.hasOwnProperty.call(MODES, format.trim().toLowerCase())
    ? MODES[format.trim().toLowerCase()]
    : undefined;
}

/** Every image field a part refers to, in first-seen order. */
export function imageFieldsIn(doc: Document): string[] {
  const seen = new Set<string>();
  for (const paragraph of elements(doc, A_NS, "p")) {
    const joined = elements(paragraph, A_NS, "t")
      .map((t) => t.textContent ?? "")
      .join("");
    for (const hit of joined.matchAll(new RegExp(FIELD.source, FIELD.flags))) {
      if (hit[1] && imageMode(hit[2])) seen.add(hit[1]);
    }
  }
  return [...seen];
}

/** What a row's image field resolves to: the bytes, or nothing. */
/**
 * `Photos\ada.PNG` and `ada.png` are the same picture.
 *
 * A cell names a file and a file picker hands back a name, and the two disagree
 * about case and about the folders in front — a user should not have to know
 * that. Directory separators BOTH ways, because a spreadsheet exported on
 * Windows writes backslashes and one exported anywhere else does not.
 *
 * Exported, and deliberately the only copy. The engine matched with one of
 * these and the pane matched with a byte-identical private one, so the screen
 * saying "All 3 pictures matched" and the merge deciding what to fill were two
 * implementations of one rule — free to drift into a pane that promises a
 * picture the merge then leaves out. The pane asks the engine everywhere else;
 * this is the same seam.
 */
export function baseName(path: string): string {
  const cut = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
  return (cut < 0 ? path : path.slice(cut + 1)).toLowerCase();
}

export type ResolveImage = (name: string) => Uint8Array | undefined;

export interface ImageOutcome {
  /** Pictures actually placed. */
  placed: number;
  /** Fields whose value named a file nobody supplied, or which had no value. */
  missing: string[];
  /** Fields whose bytes are not an image this engine can read. */
  unreadable: string[];
  /**
   * Fields placed as a plain STRETCH because no ratio could be had.
   *
   * A shape that inherits its size from a layout placeholder states no box, so
   * cover and contain cannot be computed. Reported rather than silent: a
   * stretched photo reads as a broken image, not as a fact about the template.
   */
  stretched: string[];
}

/**
 * Media parts already written for this run, keyed by content.
 *
 * Deduplication is not a nicety here. A logo on all 240 rows is 240 copies of
 * the same bytes without it, and the package is handed to the host as one
 * base64 string — a 48 MB deck becomes 64 MB across the wire, to a host that
 * already stalls. Keyed by the BYTES rather than the file name, so the same
 * picture supplied twice under two names is still stored once.
 */
export class MediaCache {
  private readonly byContent = new Map<string, string>();

  constructor(private readonly pkg: Pkg) {}

  /**
   * The media part for these bytes, adding it the first time.
   *
   * The key is length plus a cheap rolling digest of the content. A true hash
   * would be better and needs a crypto API this engine cannot assume in every
   * runtime it has to work in; a collision here would reuse the wrong picture,
   * so the length goes in the key as well and the digest walks the WHOLE
   * buffer rather than a sample.
   */
  async part(bytes: Uint8Array, extension: string, contentType: string): Promise<string> {
    const key = `${bytes.length}:${digest(bytes)}:${extension}`;
    const already = this.byContent.get(key);
    if (already) return already;
    const path = `ppt/media/image${this.pkg.nextMediaNumber()}.${extension}`;
    this.pkg.setBytes(path, bytes);
    await this.pkg.addContentTypeDefault(extension, contentType);
    this.byContent.set(key, path);
    return path;
  }
}

/** A rolling digest over every byte. Not cryptographic — see `MediaCache.part`. */
function digest(bytes: Uint8Array): string {
  let h = 0x811c9dc5;
  for (const byte of bytes) {
    h ^= byte;
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(36);
}

/**
 * Place every image field in a part, and take its placeholder away.
 *
 * The shape is found from the TEXT NODE the placeholder sits in, so a slide
 * with four picture frames fills each from its own field rather than all four
 * from the first. A field whose text is not inside a `<p:sp>` at all — in a
 * table cell, say — is reported as missing rather than guessed at.
 */
export async function placeImages(
  pkg: Pkg,
  path: string,
  doc: Document,
  resolve: ResolveImage,
  media: MediaCache,
): Promise<ImageOutcome> {
  const out: ImageOutcome = { placed: 0, missing: [], unreadable: [], stretched: [] };
  // Snapshot the shapes first: filling one edits its `<p:spPr>`, and walking a
  // live list while editing it is how a pass silently skips every other shape.
  for (const sp of shapesIn(doc)) {
    for (const paragraph of elements(sp, A_NS, "p")) {
      const texts = elements(paragraph, A_NS, "t");
      const joined = texts.map((t) => t.textContent ?? "").join("");
      const hit = [...joined.matchAll(new RegExp(FIELD.source, FIELD.flags))].find((h) => imageMode(h[2]));
      if (!hit?.[1]) continue;
      const name = hit[1];
      const mode = imageMode(hit[2]) as FillMode;

      const bytes = resolve(name);
      if (!bytes) {
        // The placeholder STAYS. Same rule the text pass follows for a field
        // with no column: a blank frame looks finished and is not.
        if (!out.missing.includes(name)) out.missing.push(name);
        continue;
      }
      const info = readImage(bytes);
      if (!info) {
        if (!out.unreadable.includes(name)) out.unreadable.push(name);
        continue;
      }

      const target = shapeOf(texts[0] ?? paragraph);
      if (!target) {
        if (!out.missing.includes(name)) out.missing.push(name);
        continue;
      }
      const partPath = await media.part(bytes, info.extension, info.contentType);
      const rId = await pkg.addRel(path, IMAGE_REL_TYPE, relativeTo(path, partPath));
      const placed = fillShapeWithImage(doc, target, rId, mode, { w: info.width, h: info.height });
      if (placed.mode === "stretch" && mode !== "stretch" && !out.stretched.includes(name)) {
        out.stretched.push(name);
      }
      // The placeholder text goes, now that the picture is there. Left behind
      // it prints over the photo it was asking for.
      for (const t of texts) t.textContent = "";
      out.placed++;
    }
  }
  return out;
}

/** `ppt/slides/slide1.xml` + `ppt/media/image2.png` → `../media/image2.png`. */
function relativeTo(owner: string, target: string): string {
  const from = owner.split("/").slice(0, -1);
  const to = target.split("/");
  while (from.length && to.length > 1 && from[0] === to[0]) {
    from.shift();
    to.shift();
  }
  return [...from.map(() => ".."), ...to].join("/");
}
