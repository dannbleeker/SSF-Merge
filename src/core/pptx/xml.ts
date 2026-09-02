/**
 * XML for the package layer.
 *
 * One implementation everywhere, deliberately. `@xmldom/xmldom` is pure
 * JavaScript, so the task pane, the build scripts and the test suite all parse and
 * serialise with the same code. Reaching for the browser's native `DOMParser`
 * when it happens to exist would buy a little speed and a class of bug this
 * project cannot afford: a merge that works in the suite and produces a file
 * PowerPoint refuses to open, because two parsers disagreed about a namespace
 * declaration nobody was looking at.
 */
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

/** DrawingML: runs, paragraphs, text bodies. Where the placeholders live. */
export const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
/** PresentationML: slides, the slide id list, custom data lists. */
export const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
/** Relationship references *inside* a part (`r:id="rId3"`). */
export const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/** The `.rels` parts themselves, which use a different namespace from `r:id`. */
export const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
/** DrawingML charts: the plot, and the cached strings behind its labels. */
export const C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
/**
 * Markup Compatibility: `<mc:AlternateContent>` and its `Choice`/`Fallback`
 * branches, which is how a slide carries a feature older hosts cannot read.
 */
export const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
/**
 * A MODERN chart's own namespace — waterfall, funnel, treemap, sunburst and the
 * rest, which are not `<c:chartSpace>` at all.
 *
 * One namespace for the PART. The compatibility token on the slide that wraps
 * it is `cx1`, `cx2` or `cx4` against three dated namespaces, and none of those
 * appear inside the chart itself.
 */
export const CX_NS = "http://schemas.microsoft.com/office/drawing/2014/chartex";
/** SpreadsheetML, for the workbook embedded behind a chart. */
export const SSML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
/** `[Content_Types].xml`. */
export const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

/**
 * A UTF-8 byte order mark at the very start of an XML part.
 *
 * Legal in an OPC package and emitted by default: .NET's `UTF8Encoding` writes
 * one unless it is explicitly told not to, so a deck from any third-party
 * generator built on it carries one on every part it wrote. JSZip's
 * `async("string")` hands the character straight through — it decodes UTF-8 and
 * has no opinion about what the first code point means — so the mark reaches
 * the parser as content.
 */
const BOM = "\uFEFF";

/**
 * Parse a part, tolerating a leading byte order mark.
 *
 * `@xmldom/xmldom` refuses one outright: with a `U+FEFF` in front of the XML
 * declaration it reports "processing instruction at position 1 is an xml
 * declaration which is only at the start of the document" and THROWS out of
 * whatever was reading the part. On a slide that is the whole merge, from a
 * deck PowerPoint opens without a murmur — the mark is the producer's, not the
 * user's, and nothing in the pane could have told them which part carried it.
 *
 * Stripped here rather than at each reader because this is the one door every
 * part in the package comes through, and a second reader would be free to
 * forget. The mark is not written back: `serializeXml` emits the document, and
 * a part that keeps its own bytes keeps its own mark, which is equally legal.
 */
export function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(
    xml.startsWith(BOM) ? xml.slice(BOM.length) : xml,
    "text/xml",
  ) as unknown as Document;
}

export function serializeXml(doc: Document): string {
  return new XMLSerializer().serializeToString(doc as never);
}

/** Every descendant with this local name in the given namespace, in document order. */
export function elements(root: Document | Element, ns: string, local: string): Element[] {
  return Array.from(root.getElementsByTagNameNS(ns, local));
}

/** The first such descendant, or undefined. Never null, so callers can `??`. */
export function element(root: Document | Element, ns: string, local: string): Element | undefined {
  return elements(root, ns, local)[0];
}

/**
 * Direct children with this local name, never deeper.
 *
 * `elements` walks DESCENDANTS, which is right for "find every tag in this part"
 * and catastrophically wrong for "what does this element own". A slide's
 * `<p:cSld>` contains the whole shape tree, so `element(cSld, P_NS, "tags")`
 * finds a SHAPE's tag reference and calls it the slide's, and
 * `element(cSld, P_NS, "extLst")` finds one inside `<p:spTree>` and appends the
 * slide's creation id into it. Both shipped. Ask for children when the parent
 * is the point.
 */
export function children(parent: Element, ns: string, local: string): Element[] {
  const out: Element[] = [];
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    const el = n as Element;
    if (el.nodeType === 1 && el.localName === local && el.namespaceURI === ns) out.push(el);
  }
  return out;
}

/** The first direct child with this local name, or undefined. */
export function child(parent: Element, ns: string, local: string): Element | undefined {
  return children(parent, ns, local)[0];
}

/**
 * Every relationship id this part's markup names, whatever names it.
 *
 * The question is "may this relationship go", and the only safe way to ask it
 * is of the WHOLE document: any attribute in the relationship namespace is a
 * reference, and a reference means the relationship has to stay.
 *
 * Two callers ask it, and they ask the same question: `dropUnusedImageRels`,
 * deciding which image relationships a merged copy still needs, and
 * `dropInheritedTags`, deciding which of the template's tag and comment
 * relationships may go with the reference it just removed. Both got it wrong in
 * the same way before this existed, and a second reader would be free to get it
 * wrong again — the answer belongs in one place.
 *
 * This began as `a:blip/@r:embed`, which is where a picture's image sits and is
 * not the only place an image id appears. Two of them are ordinary:
 *
 * - `<asvg:svgBlip r:embed="…">`, inside the blip's own extension list. That is
 *   how PowerPoint stores an ICON — a raster fallback in the blip and the real
 *   SVG beside it, under a SECOND image relationship. Icons are everywhere in a
 *   modern deck.
 * - `<a:blip r:link="…">`, a picture LINKED rather than embedded.
 *
 * Neither is an `r:embed` on an `a:blip`, so both relationships were dropped
 * from a merged copy whose slide still referenced them — a slide naming a
 * relationship that is not there, which is precisely what PowerPoint calls a
 * damaged file. It needed a modern chart on the slide to fire, because that is
 * what runs this pass at all, but nothing about the shape is exotic.
 *
 * Reading every `r:`-namespaced attribute rather than a list of the ones known
 * today is the conservative direction: an unknown reference keeps a
 * relationship that could have gone, where a missed one breaks the file. The
 * fallback picture is still dropped, because the replacement takes its `<p:pic>`
 * out of the document first and its id is then named by nothing.
 *
 * Matched by namespace OR by prefix. `getAttributeNS` was already paired with a
 * `getAttribute("r:embed")` fallback here for the same reason: a document that
 * came out of a host may not carry the namespace where a reader expects it.
 */
export function relationshipIdsIn(doc: Document): Set<string> {
  const used = new Set<string>();
  const walk = (node: Element): void => {
    const attrs = node.attributes;
    for (let i = 0; i < (attrs?.length ?? 0); i++) {
      const attr = attrs?.item(i);
      if (!attr) continue;
      if (attr.namespaceURI === R_NS || attr.name.startsWith("r:")) used.add(attr.value);
    }
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType === 1) walk(child as Element);
    }
  };
  if (doc.documentElement) walk(doc.documentElement);
  return used;
}

/**
 * Characters no XML document may contain, in any spelling.
 *
 * XML 1.0's `Char` production excludes most of the C0 controls, the lone
 * surrogates and `FFFE`/`FFFF`, and there is no escape for them either —
 * `&#11;` is exactly as ill-formed as the raw character. So a cell carrying one
 * cannot reach a slide by any route, and nothing upstream stops it:
 * `@xmldom/xmldom` writes such a character straight through and reads it back
 * again, so the part serialises, the suite sees a perfectly good document, and
 * every gate in this repo is green. PowerPoint parses conformingly, refuses the
 * part, and calls the whole file damaged — a finished merge lost entirely, with
 * no message naming the cell that did it.
 *
 * Reachable from an ordinary paste rather than from anything exotic.
 * `CHAR(11)` is the line break WORD keeps inside a cell, and a NUL is what a
 * mis-decoded UTF-16 export leaves behind.
 *
 * **Not Excel's, and this sentence used to say it was.** Alt+Enter puts
 * `CHAR(10)` in the cell and a bare LF on the clipboard, and LF is LEGAL XML —
 * so it never reaches this set, passed every gate in this repo, and DrawingML
 * rendered it as a hard break across five parts of a merged deck until the
 * desktop round of 2026-08-31 looked at one. A wrong sentence here is what sent
 * that round's own check at the wrong character.
 *
 * The rule for a legal-but-unwanted break is `foldCellBreaks` in `resolve.ts`,
 * and the two are apart on purpose: this is about what XML cannot carry from
 * any source, that is about what a CELL may hold.
 *
 * Replaced with a SPACE rather than dropped, because the likeliest one is a
 * line break: dropping it joins two words that were separate — "AdaLovelace" —
 * and a space cannot make that mistake in the other direction.
 *
 * This is the one place the engine changes a cell without being asked, and the
 * usual answer — show what the user typed and let them see it — is not
 * available: the character has no representation in the output format at all,
 * so every option alters it and only this one leaves a file that opens.
 *
 * Astral characters are NOT touched. A well-formed surrogate pair is one code
 * point above `FFFF`, which `\p{Surrogate}` under the `u` flag does not match;
 * only an unpaired half does, and an unpaired half is already broken text.
 */
// The rule is aimed at a control character reaching a pattern by accident. Here
// they ARE the subject: this is the set XML refuses, and it cannot be written
// without naming them.
// eslint-disable-next-line no-control-regex
const XML_FORBIDDEN = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|\p{Surrogate}/gu;

/** Text a conforming XML parser will accept. See `XML_FORBIDDEN`. */
export function xmlSafe(text: string): string {
  return text.replace(XML_FORBIDDEN, " ");
}
