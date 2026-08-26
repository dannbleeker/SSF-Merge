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

/** The highest value PowerPoint accepts in `<p:sldId id="…">`; the format caps ids below 2^31. */
const MAX_SLIDE_ID = 2_147_483_647;
/** PowerPoint's own numbering starts here, and ids below it are reserved. */
const MIN_SLIDE_ID = 256;

export class Pkg {
  private readonly docs = new Map<string, Document>();

  private constructor(private readonly zip: JSZip) {}

  static async open(input: Uint8Array | ArrayBuffer | string): Promise<Pkg> {
    const zip = await JSZip.loadAsync(input as never, typeof input === "string" ? { base64: true } : undefined);
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

  /** The next free `ppt/slides/slideN.xml` number. Never reuses a gap. */
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
