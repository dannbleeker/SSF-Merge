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
import { fillShapeWithImage, shapeOf } from "../image/place.js";
import type { FillMode } from "../image/fill.js";
import type { Pkg } from "../pptx/pkg.js";
import { REL_TYPE } from "../pptx/parts.js";
import { A_NS, elements } from "../pptx/xml.js";
import { editRuns, fieldsInText, textGroups, type Edit } from "./text.js";

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

/**
 * Whether a format spec is asking for a PICTURE at all, mode known or not.
 *
 * Wider than `imageMode` on purpose, and the width is the point. A spec this
 * engine does not recognise falls through to `applyFormat`, whose documented
 * answer is to print the cell unchanged — right for a number, wrong for a
 * picture, because the cell is a FILE NAME. One transposed letter in
 * `{{Photo|image-fit}}` put `ada.png` as text in the frame that was supposed to
 * hold the portrait, on every merged slide, with nothing reporting it.
 *
 * `image-cover` is the likeliest misspelling of the three real modes, because
 * "covers" is the manual's own word for what `image` does.
 *
 * The test is on this engine's OWN namespace rather than on what the author
 * might have meant: a format named `image`-something is an image format we do
 * not have. `picture` and `photo` are deliberately not included — those are
 * guesses about English, and guessing is what the rule exists to avoid. An
 * unrecognised image format leaves its placeholder on the slide, which is what
 * a field with no column already does, so the author sees their own typo.
 */
export function asksForImage(format: string | undefined): boolean {
  return format !== undefined && format.trim().toLowerCase().startsWith("image");
}

/** Whether a format spec asks for a picture rather than formatted text. */
export function imageMode(format: string | undefined): FillMode | undefined {
  if (format === undefined) return undefined;
  return Object.prototype.hasOwnProperty.call(MODES, format.trim().toLowerCase())
    ? MODES[format.trim().toLowerCase()]
    : undefined;
}

/**
 * Every image field a part refers to, in first-seen order.
 *
 * The SAME reader `mergeDocument` uses, never a second walk of its own. This
 * had one — DrawingML paragraphs only — so it could not see the places a
 * chart's text actually lives: a category label in a `<c:strCache>`, a series
 * name written literally, a shared or inline string. A `{{Photo|image}}`
 * written there was reported by nothing and printed verbatim onto the slide.
 *
 * NOT a chart's embedded workbook, which is a separate package this never
 * opens: `workbookFields` reads that, and reads it without looking at formats.
 * A picture field in a workbook cell is still reported by nothing.
 *
 * The PREDICATE is the caller's, because the two callers ask different
 * questions. A slide's list is "which fields will be filled with a picture",
 * and `placeImages` fills only an exact `imageMode` — so a misspelled
 * `{{Photo|images}}` counted there turned off the pane's "the pictures you
 * attached will not be placed" caution and offered a file picker for a field
 * nothing would fill. A notes page's list is "which fields ASK for a picture
 * somewhere one cannot go", and there the misspelling belongs in it: the merge
 * leaves such a field visible, so the user should hear about it.
 */
export function imageFieldsIn(doc: Document, asks: (format: string | undefined) => boolean): string[] {
  const seen = new Set<string>();
  for (const group of textGroups(doc)) {
    const joined = group.map((node) => node.textContent ?? "").join("");
    for (const hit of fieldsInText(joined)) {
      if (asks(hit.format)) seen.add(hit.name);
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
  /**
   * Fields whose shape already holds a picture from an earlier field.
   *
   * A shape has ONE fill, so a template with two image placeholders in one
   * shape is asking for something that cannot be drawn. The first wins and the
   * rest are named here with their placeholders left standing, because the
   * alternative — the second quietly overwriting the first — is a slide that
   * disagrees with its own count.
   */
  crowded: string[];
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
  /**
   * The same buffer OBJECT, already answered.
   *
   * The content key is computed before the lookup, so a picture already in the
   * cache was hashed in full on every row — and the digest deliberately walks
   * every byte. A logo on 240 rows, which is the case this class exists for,
   * cost 12.3 seconds of blocking work in a task-pane WebView to produce one
   * media part; the same run answered a repeat lookup in 57 ms once the hash
   * was skipped.
   *
   * `resolveImage` hands back the same `Uint8Array` instance for every row that
   * names one file, so identity is free and needs no arithmetic. The content
   * key stays — it is what makes two NAMES for one picture share a part, which
   * identity alone cannot see — and this only decides whether the digest has to
   * run at all. Weak, so a buffer the run has finished with is collectable.
   */
  private readonly byBuffer = new WeakMap<Uint8Array, string>();

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
    // Identity first, so the repeat case never reaches the digest. Keyed on the
    // extension too, because the same bytes asked for under two extensions are
    // two parts — the content key has always said so and this must agree.
    const seen = this.byBuffer.get(bytes);
    if (seen !== undefined && seen.endsWith(`.${extension}`)) return seen;
    const key = `${bytes.length}:${digest(bytes)}:${extension}`;
    const already = this.byContent.get(key);
    if (already) {
      this.byBuffer.set(bytes, already);
      return already;
    }
    const path = `ppt/media/image${this.pkg.nextMediaNumber()}.${extension}`;
    this.pkg.setBytes(path, bytes);
    await this.pkg.addContentTypeDefault(extension, contentType);
    this.byContent.set(key, path);
    this.byBuffer.set(bytes, path);
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
 *
 * That last sentence was written before the code could do it. The walk started
 * at `<p:sp>`, so a paragraph in a table was never visited: the check for it
 * sat downstream of the loop that could not reach it and was dead code, and
 * the field was skipped in silence: no picture, and no mention of it in any
 * list the caller could look at. Walking paragraphs and asking each which
 * shape it is in makes the documented answer the real one.
 *
 * ONE picture per shape, because a shape has one fill. A second image field in
 * the same shape used to overwrite the first, count itself into `placed`, and
 * leave a media part and a relationship behind for a picture that is not on
 * the slide — a count saying two where the deck shows one. It is reported now
 * and its placeholder is left standing.
 */
export async function placeImages(
  pkg: Pkg,
  path: string,
  doc: Document,
  resolve: ResolveImage,
  media: MediaCache,
): Promise<ImageOutcome> {
  const out: ImageOutcome = { placed: 0, missing: [], unreadable: [], stretched: [], crowded: [] };
  const filled = new Set<Element>();

  for (const paragraph of elements(doc, A_NS, "p")) {
    const texts = elements(paragraph, A_NS, "t");
    const joined = texts.map((t) => t.textContent ?? "").join("");
    // Materialised before the first `await`, because the loop edits the document.
    const hits = fieldsInText(joined);
    const edits: Edit[] = [];

    for (const hit of hits) {
      const mode = imageMode(hit.format);
      const name = hit.name;
      if (!mode) continue;

      const bytes = resolve(name);
      if (!bytes) {
        // The placeholder STAYS. Same rule the text pass follows for a field
        // with no column: a blank frame looks finished and is not.
        note(out.missing, name);
        continue;
      }
      const info = readImage(bytes);
      if (!info) {
        note(out.unreadable, name);
        continue;
      }
      const target = shapeOf(texts[0] ?? paragraph);
      if (!target) {
        note(out.missing, name);
        continue;
      }
      if (filled.has(target)) {
        note(out.crowded, name);
        continue;
      }

      const partPath = await media.part(bytes, info.extension, info.contentType);
      const rId = await pkg.addRel(path, REL_TYPE.image, relativeTo(path, partPath));
      const placed = fillShapeWithImage(doc, target, rId, mode, { w: info.width, h: info.height });
      if (placed.mode === "stretch" && mode !== "stretch") note(out.stretched, name);
      filled.add(target);
      // Only the placeholder's OWN characters. Blanking every text node in the
      // paragraph took the caption, and any neighbouring field, with it.
      edits.push({ start: hit.index, end: hit.index + hit.length, value: "" });
      out.placed++;
    }

    editRuns(texts, edits);
  }
  return out;
}

/** Report a field once, however many paragraphs hit the same problem with it. */
function note(list: string[], name: string): void {
  if (!list.includes(name)) list.push(name);
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
