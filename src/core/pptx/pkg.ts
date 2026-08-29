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
/**
 * What a chart or a SmartArt may drag out of the package with it.
 *
 * Its own styling and colours, the workbook behind it, its diagram parts, its
 * pictures, a theme override. That is the whole list, and it is an allowlist
 * because the alternative — trusting the relationships in the file — let a
 * crafted deck name `ppt/presentation.xml` as something a chart owned.
 */
const OWNABLE_BY_GRAPHIC = /^ppt\/(charts|diagrams|embeddings|media|theme)\//;

/** The package's own relationships: `ppt/presentation.xml`, and docProps. */
const ROOT_RELS = "_rels/.rels";

const NOTES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
/**
 * A slide's comments, in both spellings — classic `commentN.xml` and the
 * MODERN ones PowerPoint on the web writes under a Microsoft namespace.
 */
const COMMENT_REL_TYPES = [
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
  "http://schemas.microsoft.com/office/2018/10/relationships/comments",
];

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
   * Write a BINARY part — media, so far.
   *
   * Its own method rather than an overload of `setText`, because the two
   * differ in the one way that matters: JSZip stores a string as UTF-8 text,
   * so a PNG handed to `setText` arrives at the other end re-encoded and the
   * deck opens with a broken picture on every slide. The types are what keeps
   * them apart.
   */
  setBytes(path: string, bytes: Uint8Array): void {
    this.docs.delete(path);
    this.zip.file(path, bytes);
  }

  /** The raw bytes of a part. Never for XML — see `text`, which honours edits. */
  async bytes(path: string): Promise<Uint8Array> {
    const file = this.zip.file(path);
    if (!file) throw new Error(`ssf-merge: the package has no part "${path}"`);
    return file.async("uint8array");
  }

  /**
   * Declare a whole EXTENSION's content type, the way media is normally
   * declared.
   *
   * `[Content_Types].xml` takes two kinds of entry: a `Default` per extension
   * and an `Override` per part. Media uses defaults — one line for every `.png`
   * in the package rather than one per picture — and a merge that embeds two
   * hundred photos would otherwise add two hundred Overrides to a part
   * PowerPoint parses on open.
   *
   * A default that is already there is left alone rather than replaced: a
   * template may declare `png` for its own images, and a second entry for the
   * same extension is schema-invalid.
   */
  async addContentTypeDefault(extension: string, contentType: string): Promise<void> {
    const doc = await this.doc(CONTENT_TYPES);
    const already = elements(doc, CT_NS, "Default").some(
      (d) => (d.getAttribute("Extension") ?? "").toLowerCase() === extension.toLowerCase(),
    );
    if (already) return;
    const node = doc.createElementNS(CT_NS, "Default");
    node.setAttribute("Extension", extension);
    node.setAttribute("ContentType", contentType);
    // Defaults come before Overrides in every package PowerPoint writes, and
    // the schema's own sequence is unordered — but a reader that assumes the
    // conventional order is a reader this has to survive.
    doc.documentElement.insertBefore(node, doc.documentElement.firstChild);
  }

  /** The next free `ppt/media/imageN.<ext>`, across every extension. */
  nextMediaNumber(): number {
    let max = 0;
    this.zip.forEach((path) => {
      const n = Number(/^ppt\/media\/image(\d+)\./.exec(path)?.[1] ?? 0);
      if (n > max) max = n;
    });
    return max + 1;
  }

  /**
   * A parsed part, cached. Mutating the returned document is how a part is
   * edited; `save` serialises every document handed out this way. Nothing else
   * writes the same part, so the cache cannot go stale behind a caller.
   */
  /**
   * The next free number for a family of parts named `<prefix>N<suffix>`.
   *
   * Every family gets its OWN counter, read from the package rather than
   * carried alongside it. Part names are arbitrary and the sequences drift the
   * moment anything is deleted, so a deck with one chart can perfectly well
   * keep it in `chart3.xml` — and naming a copy after the slide, or after
   * another family's count, lands on a part that is already there. `copyPart`
   * then overwrites it silently and `addContentTypeOverride` no-ops on the
   * override already present, so the package stays structurally valid while two
   * slides share one chart. That is exactly the defect this whole file's
   * `nextNotesNumber` comment records, generalised so the next family cannot
   * repeat it.
   *
   * Never reuses a gap: the highest existing number plus one.
   */
  nextNumber(prefix: string, suffix = ".xml"): number {
    let max = 0;
    const pattern = new RegExp(`^${escapeRegExp(prefix)}(\\d+)${escapeRegExp(suffix)}$`);
    this.zip.forEach((path) => {
      const n = Number(pattern.exec(path)?.[1] ?? 0);
      if (n > max) max = n;
    });
    return max + 1;
  }

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
   *
   * So do its charts and its SmartArt, and those need a check rather than a
   * rule. Both used to be SHARED with every clone, so removing the template
   * left them referenced and alive; now each copy has its own, so the
   * template's would be left in the package with nothing pointing at it — a
   * whole chart and its embedded workbook per template slide, in a file the
   * host has to swallow as one base64 string. What may NOT go is the half that
   * is still shared on purpose: a diagram's layout, quick style and colours are
   * read-only styling every copy points at, and sweeping those would leave
   * every merged slide referencing a part that is not there. `orphanedParts`
   * is that distinction, asked of the package rather than assumed.
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

    // Its notes page and its comments, if it has any. Both belong to ONE slide
    // and are unreachable once that slide is gone, so leaving either behind
    // would ship a part nothing relates to.
    //
    // Comments joined the notes here when `cloneSlide` stopped copying them: a
    // clone no longer references the template's comment part, so removing the
    // template on the way out would otherwise strand it — a part with a
    // content-type override and nothing pointing at it.
    const relsPath = Pkg.relsPathFor(slidePath);
    if (this.has(relsPath)) {
      const rels = await this.doc(relsPath);
      for (const rel of elements(rels, PKG_REL_NS, "Relationship")) {
        const type = rel.getAttribute("Type") ?? "";
        if (type !== NOTES_REL_TYPE && !COMMENT_REL_TYPES.includes(type)) continue;
        const target = rel.getAttribute("Target");
        if (!target) continue;
        const related = resolveTarget(slidePath, target);
        // The TARGET comes out of the deck, and a deck can come from anywhere.
        // `resolveTarget` honours a leading `/` and any number of `..`, so a
        // crafted notes relationship naming `/[Content_Types].xml` — or
        // reaching it with enough `..` — would have this delete the one part a
        // presentation cannot open without. The output would be a file that
        // will not open, from a deck the user only had to be sent.
        //
        // A notes page and a comment part live under `ppt/`, always. Anything
        // resolving outside it is not what this loop collects, so it is left
        // alone rather than removed.
        if (!related.startsWith("ppt/")) continue;
        await this.removePart(related);
      }
      // Read BEFORE the slide's own relationships go, because that is what
      // makes them orphans: while this part exists it is one of the referrers.
      const orphans = await this.orphanedParts(slidePath);
      await this.removePart(relsPath);
      for (const path of orphans) await this.removePart(path);
    }
    await this.removePart(slidePath);
  }

  /**
   * The chart and SmartArt parts only this slide keeps alive.
   *
   * "Only this slide" is counted rather than assumed: every `.rels` in the
   * package is read, and a part any OTHER part references is left where it is.
   * That is what separates a template's own chart — which nothing else points
   * at once its slide goes — from a diagram's layout, which every merged copy
   * points at.
   *
   * Follows one hop further out from each one, because a chart owns its
   * workbook and a diagram's model owns the drawing: those are unreachable the
   * moment their owner goes, and are the bulk of the weight.
   */
  private async orphanedParts(slidePath: string): Promise<string[]> {
    const owned: string[] = [];
    for (const part of await this.relatedParts(slidePath)) {
      if (!/^ppt\/(charts\/chart|diagrams\/data)\d+\.xml$/.test(part) || !this.has(part)) continue;
      owned.push(part);
      for (const child of await this.relatedParts(part)) {
        // The child comes from the CHART's own relationships, which come out of
        // the deck, and a deck can be sent to somebody. Without this allowlist a
        // crafted chart relationship naming `/ppt/presentation.xml` put that
        // part into `owned`, nothing else in the package referred to it — the
        // only referrer is the root `_rels/.rels`, which the referrer scan below
        // does not read — and the sweep deleted it. The merge then finished
        // without complaint and produced a file PowerPoint cannot open. Naming
        // `/[Content_Types].xml` did the same and then threw.
        //
        // The parent above is already held to an allowlist. This is the same
        // discipline one level down: a chart or a diagram owns its styling, its
        // workbook and its media, and nothing else. Anything outside these is
        // left alone — and leaving a stranded part behind is a far better
        // failure than deleting one the presentation needs.
        if (!OWNABLE_BY_GRAPHIC.test(child)) continue;
        if (this.has(child) && !owned.includes(child)) owned.push(child);
      }
    }
    if (owned.length === 0) return [];

    // Every referrer in the package except the ones going with the slide.
    const referenced = new Set<string>();
    const relsPaths: string[] = [];
    this.zip.forEach((path) => {
      // `_rels/.rels` — the package's OWN relationships — has no directory in
      // front of it, so a test for "/_rels/" misses it. It was missed, and it
      // is the only referrer of `ppt/presentation.xml` and of docProps: a part
      // named just from there was invisible to this scan and looked orphaned.
      if (!path.endsWith(".rels")) return;
      if (path.includes("/_rels/") || path === ROOT_RELS) relsPaths.push(path);
    });
    const ownerOf = (rels: string): string => {
      const i = rels.indexOf("/_rels/");
      // The root's owner is the package itself, which is not a part. The empty
      // string is what `relsPathFor` and `resolveTarget` both read as "at the
      // root", so it needs no special case beyond this one.
      if (i < 0) return "";
      return `${rels.slice(0, i)}/${rels.slice(i + 7, -".rels".length)}`;
    };
    for (const rels of relsPaths) {
      const owner = ownerOf(rels);
      if (owner === slidePath || owned.includes(owner)) continue;
      for (const target of await this.relatedParts(owner)) referenced.add(target);
    }
    return owned.filter((path) => !referenced.has(path));
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
  const slash = ownerPart.lastIndexOf("/");
  // The same root-part trap `relsPathFor` documents, and here it had a
  // consequence. `lastIndexOf` answers -1 for a part at the package root, and
  // `slice(0, -1)` then drops its last character; an empty base also splits to
  // `[""]`, which prefixes every answer with a slash. So the package's own
  // `_rels/.rels` could not be resolved at all — which is why the referrer scan
  // in `orphanedParts` skipped it, and why a part named only from there looked
  // unreferenced. `""` is the root, and it now means that.
  const base = slash < 0 ? "" : ownerPart.slice(0, slash);
  const parts = base === "" ? [] : base.split("/");
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
