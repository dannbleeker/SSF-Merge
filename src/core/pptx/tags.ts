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
import { REL_TYPE } from "./parts.js";
import { P_NS, R_NS, child, element, elements, parseXml, xmlSafe } from "./xml.js";

const TAGS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml";
const TAG_LST_NS = `xmlns:p="${P_NS}"`;

/** Tag keys are stored uppercase by PowerPoint, and some APIs require it. Ours are written that way. */
export const TAG_RUN = "SSF_MERGE_RUN";
export const TAG_BLOCK = "SSF_MERGE_BLOCK";
export const TAG_SEQ = "SSF_MERGE_SEQ";
export const TAG_RECORD = "SSF_MERGE_RECORD";
// There was a TAG_TEMPLATE here, "the text to put back when preview ends". It
// existed for a preview that wrote one row onto the real template slide through
// Office.js and restored it afterwards — a design this repo's own rejected list
// forbids in the same breath, because setting a shape's text that way
// re-authors it (office-js#5858) and restoring goes through the same API that
// did the damage. The text would come back; the formatting would not, silently,
// on the master copy every merged slide is cloned from. A preview is an
// ordinary one-row merge now, inserted and then swept, so nothing is stored.

/**
 * Escape a value for an XML attribute — including the whitespace.
 *
 * The five markup characters are the obvious half. The other half is that an
 * XML parser NORMALISES an attribute value: a literal newline, carriage return
 * or tab inside one is read back as a SPACE. Writing them literally therefore
 * loses them, and the loss happens on the first merge and looks stable
 * afterwards, which is the shape that never gets reported.
 *
 * It cannot touch our own tags — a run id and a record number have no
 * whitespace in them. It reaches a FOREIGN tag, which `mergeTagPart` carries
 * through untouched and `docs/MANUAL.md` promises survives a merge. An add-in
 * keeping anything formatted in a tag got it back on one line.
 *
 * **A third half: the characters XML cannot carry AT ALL.** Escaping was the
 * whole of this and it is not enough — `&#11;` is exactly as ill-formed as the
 * byte, so a C0 control, a lone surrogate or U+FFFE in a foreign tag produced a
 * part PowerPoint refuses, on every merged slide, reported as a damaged file
 * with nothing naming the cause. `xmlSafe` is that rule, and it is shared with
 * the slide-text writer rather than copied: this and that are the only two
 * places in the engine that build XML by concatenation, and two copies of
 * "what XML can hold" is two things to drift.
 *
 * It runs FIRST, because escaping a character that may not be written is
 * writing it.
 */
function xmlAttr(s: string): string {
  return xmlSafe(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;")
    .replace(/\n/g, "&#10;")
    .replace(/\r/g, "&#13;")
    .replace(/\t/g, "&#9;");
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
  // Through `Pkg`'s own counter, which reads the package once and then keeps
  // itself current. The walk this replaces asked the package for every number
  // from 1 up on every call, and a merge calls it once per slide — so tagging
  // 2000 slides made 2,000,000 lookups. Same contract, and `nextNumber` states
  // it in the same words: the highest in use plus one, never a gap.
  return pkg.nextNumber("ppt/tags/tag");
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

  // The slide's OWN reference. `element` walks descendants, so this used to
  // find a <p:tags> inside a shape's <p:nvPr> and append the slide's merge
  // metadata to that shape's tag part — leaving the slide with no slide-level
  // tags at all, which is the exact read undo depends on.
  const custData = child(cSld, P_NS, "custDataLst");
  const existing = custData ? child(custData, P_NS, "tags") : undefined;
  if (existing) {
    const rId = existing.getAttributeNS(R_NS, "id") ?? existing.getAttribute("r:id");
    const target = rId ? await pkg.relTarget(slidePath, rId) : undefined;
    // `pkg.has`, not just a resolved target. `relTarget` answers what the
    // relationship POINTS AT, and a relationship can point at a part that is
    // not in the package — a deck another tool wrote, or one PowerPoint
    // repaired by dropping the part and leaving the reference. `pkg.text`
    // throws by name for a missing part, so a slide like that killed the whole
    // merge here while `readSlideTags` two functions down guarded with exactly
    // this test and returned an empty map. A reader that degrades and a writer
    // that throws on the same markup is the pair worth never shipping.
    if (target && pkg.has(target)) {
      pkg.setText(target, mergeTagPart(await pkg.text(target), entries));
      return;
    }
    // The reference is there and leads nowhere. It has to GO before a fresh
    // one is written, because `CT_CustomerDataList` allows at most one
    // `<p:tags>` child and the fall-through below appends into this same
    // `<p:custDataLst>` — so leaving it produced two, which is schema-invalid,
    // and `readSlideTags` reads the FIRST. The run's own tag was then
    // invisible to every reader of it: the pane could not report the slides it
    // had made and undo could not find them to take back, on a deck that
    // opened perfectly well.
    //
    // The dangling RELATIONSHIP is deliberately left alone. It was in the deck
    // before this ran and removing relationships is the operation that has
    // twice produced damage here — an id freed by a delete is handed to the
    // next thing that asks for one — so this writer repairs what it is
    // responsible for and nothing else.
    existing.parentNode?.removeChild(existing);
  }

  const n = nextTagNumber(pkg);
  const part = `ppt/tags/tag${n}.xml`;
  pkg.setText(part, tagPartXml(entries));
  await pkg.addContentTypeOverride(`/${part}`, TAGS_CONTENT_TYPE);
  const rId = await pkg.addRel(slidePath, REL_TYPE.tags, `../tags/tag${n}.xml`);

  // `CT_CommonSlideData` orders its children `bg?, spTree, custDataLst?,
  // controls?, extLst?`, so the list goes immediately after the shape tree.
  const tags = doc.createElementNS(P_NS, "p:tags");
  tags.setAttributeNS(R_NS, "r:id", rId);
  const already = custData;
  if (already) {
    // CT_CustomerDataList allows one <p:tags>, and by here there is none —
    // either the list never had one, or the one that led nowhere was taken out
    // above. A list holding only <p:custData> children is legal and common.
    already.appendChild(tags);
    return;
  }
  const custDataLst = doc.createElementNS(P_NS, "p:custDataLst");
  custDataLst.appendChild(tags);
  const spTree = child(cSld, P_NS, "spTree");
  if (!spTree) throw new Error(`ssf-merge: ${slidePath} has no <p:spTree>`);
  spTree.parentNode?.insertBefore(custDataLst, spTree.nextSibling);
}

/**
 * Append or replace entries in an existing tag part, keeping the ones we do not
 * own.
 *
 * Parsed, not pattern-matched. The regex this replaces read attribute VALUES as
 * raw source, so `val="Ben &amp; Jerry"` came back with the entity intact and
 * was escaped a second time on write: one merge turned it into `Ben &amp;amp;
 * Jerry`, two into `Ben &amp;amp;amp; Jerry`, and a reader saw the literal
 * `&amp;` on screen. It also insisted on one attribute order and a self-closing
 * tag, so PowerPoint's own perfectly legal spellings — `val` before `name`,
 * single quotes, a separate closing tag — matched nothing and the foreign tag
 * was DROPPED. `docs/MANUAL.md` promises those survive.
 *
 * The parser decodes; `tagPartXml` encodes exactly once. That round trip is
 * what makes repeated merges stable.
 */
export function mergeTagPart(xml: string, entries: [string, string][]): string {
  const kept: [string, string][] = [];
  const incoming = new Set(entries.map(([k]) => k));
  for (const tag of elements(parseXml(xml), P_NS, "tag")) {
    const name = tag.getAttribute("name");
    if (name && !incoming.has(name)) kept.push([name, tag.getAttribute("val") ?? ""]);
  }
  return tagPartXml([...kept, ...entries]);
}

/** Read a slide's tags, for the pane and for tests. */
export async function readSlideTags(pkg: Pkg, slidePath: string): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  const doc = await pkg.doc(slidePath);
  const cSld = element(doc, P_NS, "cSld");
  const custData = cSld ? child(cSld, P_NS, "custDataLst") : undefined;
  const ref = custData ? child(custData, P_NS, "tags") : undefined;
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
