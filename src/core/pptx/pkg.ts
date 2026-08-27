/**
 * A .pptx as what it is: a zip of XML parts that reference each other.
 *
 * Everything the merge engine does happens here rather than through Office.js,
 * for reasons that are recorded rather than assumed. A sibling add-in drawing
 * charts shape by shape logged a 680-second run that shipped duplicate slides;
 * the same product's one-call deck insert has none of those failure surfaces.
 * A file handed to PowerPoint as base64 is one call to lose, and a slide-count
 * delta proves whether it landed.
 */
import JSZip from "jszip";
import { CT_NS, PKG_REL_NS, P_NS, R_NS, element, elements, parseXml, serializeXml } from "./xml.js";

const CONTENT_TYPES = "[Content_Types].xml";
const PRESENTATION = "ppt/presentation.xml";
const NOTES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";

/** The highest value PowerPoint accepts in `<p:sldId id="…">`; the format caps ids below 2^31. */
const MAX_SLIDE_ID = 2_147_483_647;
/** PowerPoint's own numbering starts here, and ids below it are reserved. */
const MIN_SLIDE_ID = 256;

export class Pkg {
  private readonly docs = new Map<string, Document>();

  private constructor(private readonly zip: JSZip) {}

  static async open(input: Uint8Array | ArrayBuffer | string): Promise<Pkg> {
    const zip = await JSZip.loadAsync(input, typeof input === "string" ? { base64: true } : undefined);
    return new Pkg(zip);
  }

  /**
   * Raw text of a part, including edits not yet written back to the zip.
   *
   * A parsed document handed out by `doc` is the live copy of that part, so
   * reading the text straight off the zip would answer with the version before
   * the edit. Every caller would then have to know which parts had been touched,
   * which is exactly the kind of bookkeeping that goes wrong once and is wrong
   * silently: the merge would look right in memory and ship the template.
   */
  async text(path: string): Promise<string> {
    const cached = this.docs.get(path);
    if (cached) return serializeXml(cached);
    const file = this.zip.file(path);
    if (!file) throw new Error(`ssf-merge: the package has no part "${path}"`);
    return file.async("string");
  }

  async maybeText(path: string): Promise<string | undefined> {
    const cached = this.docs.get(path);
    if (cached) return serializeXml(cached);
    return this.zip.file(path)?.async("string");
  }

  has(path: string): boolean {
    return this.zip.file(path) !== null;
  }

  setText(path: string, xml: string): void {
    this.docs.delete(path);
    this.zip.file(path, xml);
  }

  /**
   * A parsed part, cached. Mutating the returned document is how a part is
   * edited; `save` serialises every document handed out this way. Nothing else
   * writes the same part, so the cache cannot go stale behind a caller.
   */
  async doc(path: string): Promise<Document> {
    const cached = this.docs.get(path);
    if (cached) return cached;
    const doc = parseXml(await this.text(path));
    this.docs.set(path, doc);
    return doc;
  }

  /**
   * Copy a part verbatim. Used by slide cloning for rels and for notes pages.
   *
   * An edited source is written back first. Cloning a slide whose text had
   * already been merged would otherwise copy the version from disk, and the
   * copy would silently carry the placeholders instead of the values.
   */
  async copyPart(from: string, to: string): Promise<void> {
    const pending = this.docs.get(from);
    if (pending) this.zip.file(from, serializeXml(pending));
    const file = this.zip.file(from);
    if (!file) throw new Error(`ssf-merge: cannot copy "${from}", it is not in the package`);
    this.zip.file(to, await file.async("uint8array"));
    this.docs.delete(to);
  }

  // ---- relationships -------------------------------------------------------

  /** `ppt/slides/slide1.xml` → `ppt/slides/_rels/slide1.xml.rels`. */
  static relsPathFor(part: string): string {
    const slash = part.lastIndexOf("/");
    // A part at the package ROOT has no directory, and `lastIndexOf` answers
    // -1 for it: `slice(0, -1)` then drops the part's last CHARACTER and the
    // result is a plausible-looking path to nowhere —
    // `[Content_Types].xm/_rels/[Content_Types].xml.rels`. Nothing calls this
    // with a root part today, so it has never bitten; it would fail silently
    // when something did, which is the kind worth closing on sight.
    if (slash < 0) return `_rels/${part}.rels`;
    return `${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`;
  }

  /**
   * Add a relationship to a part and return its new `rId`.
   *
   * The id is the highest existing number plus one rather than the count, so a
   * package whose relationships were never renumbered after a deletion cannot
   * produce a duplicate.
   */
  async addRel(ownerPart: string, type: string, target: string): Promise<string> {
    const path = Pkg.relsPathFor(ownerPart);
    if (!this.has(path)) {
      this.setText(
        path,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="${PKG_REL_NS}"/>`,
      );
    }
    const doc = await this.doc(path);
    const root = doc.documentElement;
    let max = 0;
    for (const rel of elements(doc, PKG_REL_NS, "Relationship")) {
      const n = Number(/^rId(\d+)$/.exec(rel.getAttribute("Id") ?? "")?.[1] ?? 0);
      if (n > max) max = n;
    }
    const id = `rId${max + 1}`;
    const rel = doc.createElementNS(PKG_REL_NS, "Relationship");
    rel.setAttribute("Id", id);
    rel.setAttribute("Type", type);
    rel.setAttribute("Target", target);
    root.appendChild(rel);
    return id;
  }

  /**
   * Every package path a part relates to, one hop out.
   *
   * External targets are skipped: a hyperlink's `Target` is a URL and
   * `resolveTarget` would answer a package path that does not exist.
   *
   * One hop is enough for what asks: a chart and a SmartArt diagram are both
   * related directly from the slide that shows them.
   */
  async relatedParts(ownerPart: string): Promise<string[]> {
    const path = Pkg.relsPathFor(ownerPart);
    if (!this.has(path)) return [];
    const doc = await this.doc(path);
    const out: string[] = [];
    for (const rel of elements(doc, PKG_REL_NS, "Relationship")) {
      if ((rel.getAttribute("TargetMode") ?? "") === "External") continue;
      const target = rel.getAttribute("Target");
      if (!target) continue;
      const resolved = resolveTarget(ownerPart, target);
      if (!out.includes(resolved)) out.push(resolved);
    }
    return out;
  }

  /** Resolve one `r:id` in a part to the package path it points at. */
  async relTarget(ownerPart: string, rId: string): Promise<string | undefined> {
    const path = Pkg.relsPathFor(ownerPart);
    if (!this.has(path)) return undefined;
    const doc = await this.doc(path);
    const rel = elements(doc, PKG_REL_NS, "Relationship").find((r) => r.getAttribute("Id") === rId);
    const target = rel?.getAttribute("Target");
    if (!target) return undefined;
    return resolveTarget(ownerPart, target);
  }

  // ---- content types -------------------------------------------------------

  /**
   * Declare a part's content type. Without this the file opens as damaged, and
   * PowerPoint does not say which part it could not classify.
   */
  async addContentTypeOverride(partName: string, contentType: string): Promise<void> {
    const doc = await this.doc(CONTENT_TYPES);
    const already = elements(doc, CT_NS, "Override").some((o) => o.getAttribute("PartName") === partName);
    if (already) return;
    const override = doc.createElementNS(CT_NS, "Override");
    override.setAttribute("PartName", partName);
    override.setAttribute("ContentType", contentType);
    doc.documentElement.appendChild(override);
  }

  // ---- slides --------------------------------------------------------------

  /** The deck's slides in presentation order, as package paths. */
  async slidePaths(): Promise<string[]> {
    const pres = await this.doc(PRESENTATION);
    const list = element(pres, P_NS, "sldIdLst");
    if (!list) return [];
    const out: string[] = [];
    for (const sldId of elements(list, P_NS, "sldId")) {
      const rId = sldId.getAttributeNS(R_NS, "id") ?? sldId.getAttribute("r:id");
      if (!rId) continue;
      const target = await this.relTarget(PRESENTATION, rId);
      if (target) out.push(target);
    }
    return out;
  }

  /**
   * Take a slide out of the deck entirely.
   *
   * Written for the merge run, which produces a package holding the TEMPLATE
   * slides and the copies made from them and must hand PowerPoint only the
   * copies. Inserting the template block again would put the user's own
   * placeholder slides back into their deck, right after the merged ones, on
   * every run.
   *
   * The alternative was to insert everything and name only the copies through
   * `insertSlidesFromBase64`'s `sourceSlideIds`. That takes ids in the host's
   * own `256#3561048925` spelling, which for a package not yet in the
   * presentation would have to be CONSTRUCTED rather than read from a Slide —
   * an assumption no round in a real host has tested, and one whose failure
   * mode is `SlideNotFound` and nothing inserted. Removing the slides here is
   * ours to get right and the suite can check it.
   *
   * Five things reference a slide and all five go: the id list entry, the
   * presentation relationship, the content-type override, its own
   * relationships, and the part. A notes page belongs to exactly one slide, so
   * it goes with it.
   */
  async removeSlide(slidePath: string): Promise<void> {
    const pres = await this.doc(PRESENTATION);
    const list = element(pres, P_NS, "sldIdLst");
    for (const sldId of list ? elements(list, P_NS, "sldId") : []) {
      const rId = sldId.getAttributeNS(R_NS, "id") ?? sldId.getAttribute("r:id");
      if (!rId) continue;
      if ((await this.relTarget(PRESENTATION, rId)) !== slidePath) continue;
      sldId.parentNode?.removeChild(sldId);
      const rels = await this.doc(Pkg.relsPathFor(PRESENTATION));
      for (const rel of elements(rels, PKG_REL_NS, "Relationship")) {
        if (rel.getAttribute("Id") === rId) rel.parentNode?.removeChild(rel);
      }
    }

    // Its notes page, if it has one. A notes slide belongs to one slide and is
    // unreachable once that slide is gone, so leaving it behind would ship a
    // part nothing relates to.
    const relsPath = Pkg.relsPathFor(slidePath);
    if (this.has(relsPath)) {
      const rels = await this.doc(relsPath);
      for (const rel of elements(rels, PKG_REL_NS, "Relationship")) {
        if (rel.getAttribute("Type") !== NOTES_REL_TYPE) continue;
        const target = rel.getAttribute("Target");
        if (target) await this.removePart(resolveTarget(slidePath, target));
      }
      await this.removePart(relsPath);
    }
    await this.removePart(slidePath);
  }

  /** Drop a part, its own relationships and its content-type override. */
  private async removePart(path: string): Promise<void> {
    const relsPath = Pkg.relsPathFor(path);
    if (this.has(relsPath)) {
      this.docs.delete(relsPath);
      this.zip.remove(relsPath);
    }
    const types = await this.doc(CONTENT_TYPES);
    for (const override of elements(types, CT_NS, "Override")) {
      if (override.getAttribute("PartName") === `/${path}`) override.parentNode?.removeChild(override);
    }
    this.docs.delete(path);
    this.zip.remove(path);
  }

  /** The next free `ppt/slides/slideN.xml` number. Never reuses a gap. */
  /**
   * The next free `ppt/notesSlides/notesSlideN.xml` number.
   *
   * Its OWN counter, not the slide's. Part names in a package are arbitrary and
   * the two sequences drift apart the moment a slide is deleted, so a deck with
   * one slide can perfectly well keep its notes in `notesSlide2.xml`. Naming a
   * clone's notes after the slide number then lands on a part that is already
   * there — and `copyPart` overwrites silently while `addContentTypeOverride`
   * no-ops on the override that already exists, so the package stays structurally
   * valid and is wrong in two ways at once: the clone shares the template's
   * notes page (so the NEXT clone copies notes that have already been merged,
   * and record 2's slide ships record 1's text), and removing the template on
   * the way out deletes that shared part, leaving a slide whose notes
   * relationship points at nothing.
   *
   * Both were reproduced on real bytes before this existed. `nextTagNumber` had
   * the right shape all along, one file over: ask whether the path is free
   * rather than assume it.
   */
  nextNotesNumber(): number {
    let n = 1;
    while (this.has(`ppt/notesSlides/notesSlide${n}.xml`)) n++;
    return n;
  }

  nextSlideNumber(): number {
    let max = 0;
    this.zip.forEach((path) => {
      const n = Number(/^ppt\/slides\/slide(\d+)\.xml$/.exec(path)?.[1] ?? 0);
      if (n > max) max = n;
    });
    return max + 1;
  }

  /**
   * Append a slide to the deck's own order and return the id it was given.
   *
   * The id has to be unique and inside the format's range. Taking the highest
   * in use plus one satisfies both, and appending rather than inserting is what
   * makes the merged block land after the template instead of in front of it.
   */
  async appendSldId(rId: string): Promise<number> {
    const pres = await this.doc(PRESENTATION);
    const list = element(pres, P_NS, "sldIdLst");
    if (!list) throw new Error("ssf-merge: presentation.xml has no <p:sldIdLst>");
    let max = MIN_SLIDE_ID - 1;
    for (const sldId of elements(list, P_NS, "sldId")) {
      const n = Number(sldId.getAttribute("id") ?? 0);
      if (n > max) max = n;
    }
    const id = max + 1;
    if (id > MAX_SLIDE_ID) throw new Error("ssf-merge: the deck has run out of slide ids");
    const el = pres.createElementNS(P_NS, "p:sldId");
    el.setAttribute("id", String(id));
    el.setAttributeNS(R_NS, "r:id", rId);
    list.appendChild(el);
    return id;
  }

  // ---- output --------------------------------------------------------------

  /**
   * Write a part back into the zip and drop its parsed copy.
   *
   * The cache is also the dirty-part set — `flush` writes every document handed
   * out by `doc` — so nothing ever left it, and a merge held one live xmldom
   * Document per output slide on top of JSZip's copy of the same bytes. Measured
   * at 300 clones of a 124 KB slide: 1697 MB of heap against 93 MB with the
   * documents released, and 400 records died outright under a 2 GB limit, which
   * is the size a task-pane WebView is working in.
   *
   * Releasing is behaviour-neutral: every PART is byte-identical either way,
   * which is asserted rather than assumed. The ZIP is not, and that is not a
   * difference in output — JSZip stamps an entry time whenever a file is
   * written, so any two builds differ. Compare parts, never the archive.
   *
   * Call it for a part nothing will read again; parts the
   * run keeps amending, `[Content_Types].xml` and `ppt/presentation.xml` among
   * them, must NOT be released or every clone reparses them.
   */
  release(path: string): void {
    const doc = this.docs.get(path);
    if (doc) this.setText(path, serializeXml(doc));
  }

  /**
   * How many parts are parsed and held right now.
   *
   * A diagnostic, and the only way to state the property `release` exists for:
   * that a merge's held-document count does not grow with the number of records.
   * A memory assertion would be flaky; this one is exact.
   */
  cachedParts(): number {
    return this.docs.size;
  }

  /**
   * Every part in the package, in the zip's own order.
   *
   * A measurement rather than a manipulation: the package is handed to
   * PowerPoint as base64 and then goes out of scope, so when the host answers
   * `InvalidArgument` the file that caused it no longer exists anywhere. What
   * survives has to be counted while it is still here, and "how many parts"
   * separates a package missing its content types from one that is merely
   * large.
   *
   * Directory entries are excluded — JSZip records them and they are not parts.
   */
  partNames(): string[] {
    return Object.keys(this.zip.files).filter((name) => !this.zip.files[name]?.dir);
  }

  private flush(): void {
    for (const [path, doc] of this.docs) this.zip.file(path, serializeXml(doc));
  }

  async toBase64(): Promise<string> {
    this.flush();
    return this.zip.generateAsync({ type: "base64", compression: "DEFLATE" });
  }

  async toBytes(): Promise<Uint8Array> {
    this.flush();
    return this.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }
}

/** Resolve a relationship target, which may be relative, against the part that holds it. */
export function resolveTarget(ownerPart: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const base = ownerPart.slice(0, ownerPart.lastIndexOf("/"));
  const parts = base.split("/");
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}
