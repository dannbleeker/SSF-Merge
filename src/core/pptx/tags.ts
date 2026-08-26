/**
 * PowerPoint tags, written into the file instead of asked of the API.
 *
 * This is the single most load-bearing decision in the engine. On PowerPoint
 * for the web a slide the run has just added does not round-trip through
 * `slides.getItem(id)`, and tag writes through a shape proxy are refused
 * outright: a sibling add-in logged 46 `InvalidParam passed to GetItem(id)`
 * failures in one run and needed a whole recovery pass to claw some of them
 * back. A tag written into the package before the insert cannot be refused,
 * because nothing is asked.
 *
 * The shape is `ppt/tags/tagN.xml` holding `<p:tagLst>`, related from the owner
 * part, and referenced from the owner's `<p:custDataLst><p:tags r:id="…"/>`.
 */
import { Pkg } from "./pkg.js";
import { P_NS, R_NS, element, elements } from "./xml.js";

const TAGS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags";
const TAGS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml";
const TAG_LST_NS = `xmlns:p="${P_NS}"`;

/** Tag keys are stored uppercase by PowerPoint, and some APIs require it. Ours are written that way. */
export const TAG_RUN = "SSF_MERGE_RUN";
export const TAG_BLOCK = "SSF_MERGE_BLOCK";
export const TAG_SEQ = "SSF_MERGE_SEQ";
export const TAG_RECORD = "SSF_MERGE_RECORD";
export const TAG_TEMPLATE = "SSF_MERGE_TEMPLATE";

function xmlAttr(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export function tagPartXml(entries: [string, string][]): string {
  const tags = entries.map(([name, val]) => `<p:tag name="${xmlAttr(name)}" val="${xmlAttr(val)}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:tagLst ${TAG_LST_NS}>${tags}</p:tagLst>`;
}

/**
 * The next free `ppt/tags/tagN.xml`.
 *
 * Blind use of `tag1.xml` is the trap here: a template that already carries one
 * would have it overwritten, and every slide pointing at it would silently lose
 * its tags. The number is taken from the package, never assumed.
 */
export function nextTagNumber(pkg: Pkg): number {
  let n = 1;
  while (pkg.has(`ppt/tags/tag${n}.xml`)) n++;
  return n;
}

/**
 * Attach tags to a slide, merging with whatever is already there.
 *
 * `CT_CustomerDataList` allows at most one `<p:tags>` child, so a slide that
 * already has a tag part must have its entries appended rather than a second
 * part added. A template built by another tool routinely does.
 */
export async function writeSlideTags(pkg: Pkg, slidePath: string, entries: [string, string][]): Promise<void> {
  const doc = await pkg.doc(slidePath);
  const cSld = element(doc, P_NS, "cSld");
  if (!cSld) throw new Error(`ssf-merge: ${slidePath} has no <p:cSld>`);

  const existing = element(cSld, P_NS, "tags");
  if (existing) {
    const rId = existing.getAttributeNS(R_NS, "id") ?? existing.getAttribute("r:id");
    const target = rId ? await pkg.relTarget(slidePath, rId) : undefined;
    if (target) {
      pkg.setText(target, mergeTagPart(await pkg.text(target), entries));
      return;
    }
  }

  const n = nextTagNumber(pkg);
  const part = `ppt/tags/tag${n}.xml`;
  pkg.setText(part, tagPartXml(entries));
  await pkg.addContentTypeOverride(`/${part}`, TAGS_CONTENT_TYPE);
  const rId = await pkg.addRel(slidePath, TAGS_REL_TYPE, `../tags/tag${n}.xml`);

  // `CT_CommonSlideData` orders its children `bg?, spTree, custDataLst?,
  // controls?, extLst?`, so the list goes immediately after the shape tree.
  const custDataLst = doc.createElementNS(P_NS, "p:custDataLst");
  const tags = doc.createElementNS(P_NS, "p:tags");
  tags.setAttributeNS(R_NS, "r:id", rId);
  custDataLst.appendChild(tags);
  const spTree = element(cSld, P_NS, "spTree");
  if (!spTree) throw new Error(`ssf-merge: ${slidePath} has no <p:spTree>`);
  spTree.parentNode?.insertBefore(custDataLst, spTree.nextSibling);
}

/** Append or replace entries in an existing tag part, keeping the ones we do not own. */
export function mergeTagPart(xml: string, entries: [string, string][]): string {
  const kept: [string, string][] = [];
  const re = /<p:tag\s+name="([^"]*)"\s+val="([^"]*)"\s*\/>/g;
  const incoming = new Set(entries.map(([k]) => k));
  for (const m of xml.matchAll(re)) {
    const name = m[1] ?? "";
    if (!incoming.has(name)) kept.push([name, m[2] ?? ""]);
  }
  return tagPartXml([...kept, ...entries]);
}

/** Read a slide's tags, for the pane and for tests. */
export async function readSlideTags(pkg: Pkg, slidePath: string): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  const doc = await pkg.doc(slidePath);
  const cSld = element(doc, P_NS, "cSld");
  const ref = cSld ? element(cSld, P_NS, "tags") : undefined;
  const rId = ref?.getAttributeNS(R_NS, "id") ?? ref?.getAttribute("r:id");
  if (!rId) return out;
  const target = await pkg.relTarget(slidePath, rId);
  if (!target || !pkg.has(target)) return out;
  const doc2 = await pkg.doc(target);
  for (const tag of elements(doc2, P_NS, "tag")) {
    const name = tag.getAttribute("name");
    if (name) out.set(name, tag.getAttribute("val") ?? "");
  }
  return out;
}
