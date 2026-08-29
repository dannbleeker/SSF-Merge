/**
 * XML for the package layer.
 *
 * One implementation everywhere, deliberately. `@xmldom/xmldom` is pure
 * JavaScript, so the task pane, the Node CLI and the test suite all parse and
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

export function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, "text/xml") as unknown as Document;
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
