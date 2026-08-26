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
import { Pkg } from "./pkg.js";
import { P_NS, element, elements } from "./xml.js";

const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const NOTES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
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
 */
export async function cloneSlide(pkg: Pkg, sourcePath: string, opts: CloneOptions = {}): Promise<string> {
  const n = pkg.nextSlideNumber();
  const target = `ppt/slides/slide${n}.xml`;

  await pkg.copyPart(sourcePath, target);
  const sourceRels = Pkg.relsPathFor(sourcePath);
  if (pkg.has(sourceRels)) await pkg.copyPart(sourceRels, Pkg.relsPathFor(target));

  await pkg.addContentTypeOverride(`/${target}`, SLIDE_CONTENT_TYPE);
  const rId = await pkg.addRel("ppt/presentation.xml", SLIDE_REL_TYPE, `slides/slide${n}.xml`);
  await pkg.appendSldId(rId);

  await cloneNotesSlide(pkg, target, n);
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
async function cloneNotesSlide(pkg: Pkg, slidePath: string, n: number): Promise<void> {
  const relsPath = Pkg.relsPathFor(slidePath);
  if (!pkg.has(relsPath)) return;
  const rels = await pkg.doc(relsPath);
  const notesRel = elements(rels, PKG_REL_NS, "Relationship").find((r) => r.getAttribute("Type") === NOTES_REL_TYPE);
  if (!notesRel) return;

  const oldTarget = notesRel.getAttribute("Target") ?? "";
  const oldPath = resolve(slidePath, oldTarget);
  if (!pkg.has(oldPath)) return;

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
      if (rel.getAttribute("Type") === SLIDE_REL_TYPE) rel.setAttribute("Target", `../slides/slide${n}.xml`);
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

  for (const id of Array.from(doc.getElementsByTagNameNS(P14_NS, "creationId"))) {
    id.setAttribute("val", String(value));
    return;
  }

  let extLst = element(cSld, P_NS, "extLst");
  if (!extLst) {
    extLst = doc.createElementNS(P_NS, "p:extLst");
    cSld.appendChild(extLst);
  }
  const ext = doc.createElementNS(P_NS, "p:ext");
  ext.setAttribute("uri", CREATION_ID_URI);
  const creationId = doc.createElementNS(P14_NS, "p14:creationId");
  creationId.setAttribute("xmlns:p14", P14_NS);
  creationId.setAttribute("val", String(value));
  ext.appendChild(creationId);
  extLst.appendChild(ext);
}

/** Read a slide's creation id, for tests and for the pane's diagnostics. */
export async function creationIdOf(pkg: Pkg, slidePath: string): Promise<number | undefined> {
  const doc = await pkg.doc(slidePath);
  const id = Array.from(doc.getElementsByTagNameNS(P14_NS, "creationId"))[0];
  const val = id?.getAttribute("val");
  return val === undefined || val === null ? undefined : Number(val);
}

function resolve(ownerPart: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = ownerPart.slice(0, ownerPart.lastIndexOf("/")).split("/");
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}
