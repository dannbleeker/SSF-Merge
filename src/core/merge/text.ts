/**
 * Replace placeholders without touching the formatting around them.
 *
 * The whole product rests on this file. Setting a shape's text through
 * Office.js re-authors it: office-js#5858 reports custom bullets reverting to
 * default discs, and mixed runs inside one paragraph collapse to one state. In
 * the file there is no such problem, because each `<a:r>` keeps its own
 * `<a:rPr>` and only the characters change.
 *
 * The complication is that PowerPoint splits runs wherever it likes.
 * `{{FirstName}}` is routinely stored as `{{Fir` + `stName}}` after an edit or a
 * spellcheck pass, so a per-node search finds nothing and silently merges
 * nothing. Every match here is computed against the paragraph's joined text.
 */
import { A_NS, elements } from "../pptx/xml.js";

/** `{{Field}}` or `{{Field|format}}`. Field names are word characters and dots, so `{{a.b}}` works. */
export const FIELD = /\{\{\s*([\w.]+)\s*(?:\|\s*([^{}]+?)\s*)?\}\}/g;

/** Answers the value for a field, or null to leave the placeholder visible. */
export type Resolve = (name: string, format?: string) => string | null;

interface Span {
  node: Element;
  from: number;
  to: number;
}

function spansOf(paragraph: Element): { spans: Span[]; joined: string } {
  const spans: Span[] = [];
  let joined = "";
  for (const node of elements(paragraph, A_NS, "t")) {
    const s = node.textContent ?? "";
    spans.push({ node, from: joined.length, to: joined.length + s.length });
    joined += s;
  }
  return { spans, joined };
}

/**
 * Merge one `<a:p>`. Returns true if anything changed.
 *
 * The replacement lands in the FIRST run the match touches and the covered
 * characters are removed from the rest, so the value inherits the run
 * properties of the run the placeholder started in. That is the contract a
 * template author works against: format the opening brace the way you want the
 * value to look.
 */
export function mergeParagraph(paragraph: Element, resolve: Resolve): boolean {
  const { spans, joined } = spansOf(paragraph);
  if (!spans.length) return false;

  const hits = [...joined.matchAll(FIELD)];
  if (!hits.length) return false;

  // Per-character buffers, so every offset keeps referring to the ORIGINAL
  // joined text however many fields one paragraph holds.
  const bufs = spans.map((s) => (s.node.textContent ?? "").split(""));
  let changed = false;

  for (const hit of hits) {
    const value = resolve(hit[1] ?? "", hit[2]);
    // A field nobody can resolve stays visible. Blanking it hides the author's
    // typo behind 240 slides that look finished and are not.
    if (value === null) continue;

    const start = hit.index ?? 0;
    const end = start + hit[0].length;
    let carried = false;

    spans.forEach((s, i) => {
      if (s.to <= start || s.from >= end) return;
      const buf = bufs[i];
      if (!buf) return;
      const lo = Math.max(start, s.from) - s.from;
      const hi = Math.min(end, s.to) - s.from;
      for (let k = lo; k < hi; k++) buf[k] = "";
      if (!carried) {
        buf[lo] = value;
        carried = true;
      }
    });
    changed = true;
  }
  if (!changed) return false;

  spans.forEach((s, i) => {
    const text = (bufs[i] ?? []).join("");
    s.node.textContent = text;
    // Without this PowerPoint eats a leading or trailing space, and
    // "Dear Ada ," is a support ticket.
    if (/^\s|\s$/.test(text)) s.node.setAttribute("xml:space", "preserve");
  });
  return true;
}

/**
 * Merge every paragraph in a part.
 *
 * One `getElementsByTagNameNS` over the whole document reaches shape bodies,
 * table cells and grouped shapes at any depth, because they all hold ordinary
 * `<a:p>` elements. Charts and SmartArt keep their text in separate parts and
 * are deliberately not covered here.
 */
export function mergeDocument(doc: Document, resolve: Resolve): number {
  let n = 0;
  for (const paragraph of elements(doc, A_NS, "p")) if (mergeParagraph(paragraph, resolve)) n++;
  return n;
}

/** Every field name a part refers to, in first-seen order. For the pane's field list. */
export function fieldsIn(doc: Document): string[] {
  const seen = new Set<string>();
  for (const paragraph of elements(doc, A_NS, "p")) {
    const { joined } = spansOf(paragraph);
    for (const hit of joined.matchAll(FIELD)) if (hit[1]) seen.add(hit[1]);
  }
  return [...seen];
}
