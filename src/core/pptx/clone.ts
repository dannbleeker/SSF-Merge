/**
 * Clone a slide inside the package.
 *
 * Six things have to agree or PowerPoint reports the deck as damaged, and it
 * does that by refusing to open the file without saying which one was wrong:
 * the part, its relationships, the content-type override, the presentation
 * relationship, the slide id list, and the creation id.
 *
 * The creation id is the one that is easy to miss and expensive to get wrong.
 * Office.js reports a slide as `256#3561048925`, which is `<p:sldId id>` joined
 * to `<p14:creationId val>` from the slide's own `extLst`. Re-inserting a slide
 * that carries a creation id already in the deck asks the host to hold one
 * identity twice, and office-js#6105 reports exactly that failing with
 * `InvalidArgument` on Windows desktop. Every copy gets a fresh one.
 */
import { Pkg, resolveTarget as resolve } from "./pkg.js";
import { P_NS, child, element, elements, relationshipIdsIn } from "./xml.js";
import { COMMENT_REL_TYPES, REL_TYPE } from "./parts.js";

const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main";
/** The extension slot PowerPoint keeps a slide's creation id in. */
const CREATION_ID_URI = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}";

export interface CloneOptions {
  /** Injectable for tests. Real runs want a fresh value per copy and nothing else. */
  creationId?: () => number;
}

function randomCreationId(): number {
  return Math.floor(Math.random() * 0xffff_ffff) + 1;
}

/**
 * Copy one slide and return the new part's path.
 *
 * The layout, the master, media and hyperlinks stay shared: they are read-only
 * as far as a merge is concerned, and duplicating them would bloat the deck for
 * nothing. A notes slide is the exception, because it is per-slide content and
 * a shared one means two slides editing the same notes page.
 *
 * COMMENTS are dropped rather than shared or copied, and the two routes are why.
 * A comment hangs off the slide, so the wholesale rels copy above hands every
 * clone a relationship to the TEMPLATE's comment part — measured: three slides,
 * one `modernComment_101_AEAB9DA1.xml`. A reviewer's "check this with Legal"
 * then appears on all 240 merged slides, as one shared thread.
 *
 * Copying them per clone would be worse, not better: it is the same note 240
 * times, deliberately. A comment is an annotation about the template, not
 * content the template produces.
 *
 * And dropping them is what makes the two template routes AGREE. On a 1.10 host
 * `exportAsBase64Presentation` drops comments and `ppt/authors.xml` outright —
 * office-js#6867, measured on this host on 2026-08-28, four comment parts in and
 * none out — so the subset route was already producing comment-free clones while
 * the file route produced shared ones. Two routes, two different decks, from the
 * same template.
 */
export async function cloneSlide(pkg: Pkg, sourcePath: string, opts: CloneOptions = {}): Promise<string> {
  const n = pkg.nextSlideNumber();
  const target = `ppt/slides/slide${n}.xml`;

  await pkg.copyPart(sourcePath, target);
  const sourceRels = Pkg.relsPathFor(sourcePath);
  if (pkg.has(sourceRels)) await pkg.copyPart(sourceRels, Pkg.relsPathFor(target));

  await pkg.addContentTypeOverride(`/${target}`, SLIDE_CONTENT_TYPE);
  const rId = await pkg.addRel("ppt/presentation.xml", REL_TYPE.slide, `slides/slide${n}.xml`);
  await pkg.appendSldId(rId);

  await cloneNotesSlide(pkg, target, n);
  await dropInheritedTags(pkg, target);
  await setCreationId(pkg, target, (opts.creationId ?? randomCreationId)());

  return target;
}

/**
 * Give the copy its own notes page.
 *
 * The notes slide also points back at the slide it belongs to, so the copy's
 * back-reference is repointed too. Left alone it names the template, and
 * PowerPoint then shows one slide's notes on another.
 */
async function cloneNotesSlide(pkg: Pkg, slidePath: string, slideNumber: number): Promise<void> {
  const relsPath = Pkg.relsPathFor(slidePath);
  if (!pkg.has(relsPath)) return;
  const rels = await pkg.doc(relsPath);
  const notesRel = elements(rels, PKG_REL_NS, "Relationship").find(
    (r) => r.getAttribute("Type") === REL_TYPE.notesSlide,
  );
  if (!notesRel) return;

  const oldTarget = notesRel.getAttribute("Target") ?? "";
  const oldPath = resolve(slidePath, oldTarget);
  if (!pkg.has(oldPath)) return;

  // Numbered from the NOTES parts, never from the slide. The two sequences are
  // independent — a deck with one slide can keep its notes in `notesSlide2.xml`
  // — and taking the slide's number lands on a part that is already there. See
  // `nextNotesNumber`, which records what that cost.
  const n = pkg.nextNotesNumber();
  const newPath = `ppt/notesSlides/notesSlide${n}.xml`;
  await pkg.copyPart(oldPath, newPath);
  await pkg.addContentTypeOverride(
    `/${newPath}`,
    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
  );
  notesRel.setAttribute("Target", `../notesSlides/notesSlide${n}.xml`);

  const oldNotesRels = Pkg.relsPathFor(oldPath);
  if (pkg.has(oldNotesRels)) {
    await pkg.copyPart(oldNotesRels, Pkg.relsPathFor(newPath));
    const notesRels = await pkg.doc(Pkg.relsPathFor(newPath));
    for (const rel of elements(notesRels, PKG_REL_NS, "Relationship")) {
      if (rel.getAttribute("Type") === REL_TYPE.slide) rel.setAttribute("Target", `../slides/slide${slideNumber}.xml`);
    }
  }
}

/**
 * Stamp a slide's creation id, adding the extension slot if the template has none.
 *
 * A slide with no creation id at all is legal and PowerPoint invents one on
 * open, which would make two copies indistinguishable at exactly the wrong
 * moment. Writing one is cheaper than finding out.
 */
export async function setCreationId(pkg: Pkg, slidePath: string, value: number): Promise<void> {
  const doc = await pkg.doc(slidePath);
  const cSld = element(doc, P_NS, "cSld");
  if (!cSld) throw new Error(`ssf-merge: ${slidePath} has no <p:cSld>`);

  // The slide's OWN extLst, not the first creation id anywhere in the part.
  //
  // The append path below is already scoped to `cSld` and says why: a slide
  // whose shape tree ends in its own `<p:extLst>` had the id appended THERE,
  // where PowerPoint does not look, so it invented one on open and two copies
  // were indistinguishable. This search was not scoped, and reached the same
  // failure from the other side — it found a stray id inside the shape tree,
  // updated THAT, and returned, leaving the slide with no id of its own while
  // `creationIdOf` reported the stamp had worked.
  //
  // Not hypothetical: the comment below records that an older version of this
  // function put ids exactly there, so a deck merged by it carries one, and
  // using that deck as a template is an ordinary thing to do.
  const own = child(cSld, P_NS, "extLst");
  const existing = own ? Array.from(own.getElementsByTagNameNS(P14_NS, "creationId"))[0] : undefined;
  if (existing) {
    existing.setAttribute("val", String(value));
    return;
  }

  // A direct child of cSld. `element` walks descendants, so a slide whose shape
  // tree ends in its own <p:extLst> had the creation id appended THERE, where
  // PowerPoint does not look for one — so it invented an id on open and two
  // copies were indistinguishable again, which is the collision this file
  // exists to avoid.
  let extLst = child(cSld, P_NS, "extLst");
  if (!extLst) {
    extLst = doc.createElementNS(P_NS, "p:extLst");
    cSld.appendChild(extLst);
  }
  const ext = doc.createElementNS(P_NS, "p:ext");
  ext.setAttribute("uri", CREATION_ID_URI);
  // No manual xmlns:p14. `createElementNS` binds the prefix and the serializer
  // emits the declaration itself, so setting it by hand produced
  // `<p14:creationId xmlns:p14="…" val="…" xmlns:p14="…"/>` — a duplicate
  // attribute, which XML forbids outright (WFC: Unique Att Spec). PowerPoint
  // rejects the whole package for it, and says nothing about which part.
  const creationId = doc.createElementNS(P14_NS, "p14:creationId");
  creationId.setAttribute("val", String(value));
  ext.appendChild(creationId);
  extLst.appendChild(ext);
}

/** Read a slide's creation id, for tests and for the pane's diagnostics. */
export async function creationIdOf(pkg: Pkg, slidePath: string): Promise<number | undefined> {
  const doc = await pkg.doc(slidePath);
  // Scoped like the write. Reading the first one in the part reported a stray
  // id from inside the shape tree as the slide's own, which is a diagnostic
  // agreeing with a stamp that never landed.
  const cSld = element(doc, P_NS, "cSld");
  const own = cSld ? child(cSld, P_NS, "extLst") : undefined;
  const id = own ? Array.from(own.getElementsByTagNameNS(P14_NS, "creationId"))[0] : undefined;
  const val = id?.getAttribute("val");
  return val === undefined || val === null ? undefined : Number(val);
}

/**
 * A copy starts with no tags of its own.
 *
 * The .rels are copied verbatim, so a template that already carries a tag part
 * hands the clone a relationship pointing at the TEMPLATE's `ppt/tags/tagN.xml`
 * — and `writeSlideTags` then appends the copy's metadata there, because from
 * its side the slide plainly has a tag reference. Every merged slide ends up
 * sharing one part, so all but the last record's tags are overwritten, and the
 * user's own template is stamped as merge output and matched by undo.
 *
 * A template with tags is not exotic: `docs/MANUAL.md` records BLOCK and SEQ as
 * living on "a template or merged slide", so a block picked out of an earlier
 * merge has them, and any other add-in's tags do the same.
 *
 * The relationship and the reference both go. The template's own tag part is
 * left exactly as it was — it belongs to the template.
 */
async function dropInheritedTags(pkg: Pkg, slidePath: string): Promise<void> {
  const doc = await pkg.doc(slidePath);
  const cSld = element(doc, P_NS, "cSld");
  const custData = cSld ? child(cSld, P_NS, "custDataLst") : undefined;
  const tags = custData ? child(custData, P_NS, "tags") : undefined;
  if (custData && tags) {
    custData.removeChild(tags);
    // An empty <p:custDataLst> is schema-invalid: it requires at least one
    // child. Drop the list when the tag reference was all it held.
    if (!custData.firstChild) custData.parentNode?.removeChild(custData);
  }

  const relsPath = Pkg.relsPathFor(slidePath);
  if (!pkg.has(relsPath)) return;
  const rels = await pkg.doc(relsPath);
  // Read AFTER the slide's own reference has gone, so its id is named by
  // nothing and its relationship goes with it — which is the whole point of
  // this function.
  //
  // What the id list protects is a reference somewhere ELSE. A tag part can be
  // named by a SHAPE, from `<p:nvPr><p:custDataLst><p:tags r:id="…"/>`, and
  // that is where an add-in puts its own bookkeeping: a deck touched by
  // think-cell has exactly this on the shape it hides in every slide. Removing
  // every tag relationship left that shape naming one that was gone.
  //
  // Which was not even the visible half. Deleting a relationship FREES ITS ID,
  // and `writeSlideTags` takes the next free one for this run's own BLOCK and
  // SEQ tags — so the vendor's shape came out of the merge pointing at SSF
  // Merge's merge metadata. A reference that still resolves, to somebody else's
  // data.
  const named = relationshipIdsIn(doc);
  for (const rel of elements(rels, PKG_REL_NS, "Relationship")) {
    const type = rel.getAttribute("Type") ?? "";
    if (type !== REL_TYPE.tags && !COMMENT_REL_TYPES.includes(type)) continue;
    const id = rel.getAttribute("Id") ?? "";
    if (named.has(id)) continue;
    rel.parentNode?.removeChild(rel);
  }
}

/**
 * The notes page a slide owns, if it has one.
 *
 * Exported because the merge has to reach it: a copy gets its own notes slide
 * precisely so the copies can differ, and that only pays if the placeholders in
 * it are merged too.
 */
export async function notesPathFor(pkg: Pkg, slidePath: string): Promise<string | undefined> {
  const relsPath = Pkg.relsPathFor(slidePath);
  if (!pkg.has(relsPath)) return undefined;
  const rels = await pkg.doc(relsPath);
  const rel = elements(rels, PKG_REL_NS, "Relationship").find((r) => r.getAttribute("Type") === REL_TYPE.notesSlide);
  const target = rel?.getAttribute("Target");
  if (!target) return undefined;
  const path = resolve(slidePath, target);
  return pkg.has(path) ? path : undefined;
}
