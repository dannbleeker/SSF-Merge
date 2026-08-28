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
import { A_NS, C_NS, SSML_NS, children, elements } from "../pptx/xml.js";

/**
 * `{{Field}}` or `{{Field|format}}`.
 *
 * A field name is any letter, mark, digit, underscore or dot — Unicode, not
 * ASCII. `[\w.]` matched none of `ø`, `å`, `é`, `ü`, Greek or Cyrillic, so
 * `{{Beløb}}` and `{{Måned}}` were INVISIBLE to the engine: `fieldsIn` never
 * reported them, so the pane could not flag them as unmatched either, and the
 * literal braces were printed on every merged slide. On a product whose first
 * users write Danish, that is most of a template.
 *
 * `\p{L}\p{M}\p{N}` rather than `\w` with the `u` flag, because `\w` stays
 * ASCII-only under `u`; the flag alone would have changed nothing. A column
 * name arrives verbatim from the header and the resolver already answers for
 * it, so the pattern was the only thing in the way.
 *
 * **A name may contain SPACES, and refusing them refused most real data.** The
 * class above was a list of allowed characters and a space was not on it, so
 * `{{Row Labels}}`, `{{Min. of cost}}` and `{{Sum of quantity monthly}}` — the
 * literal default headers of an Excel pivot table, which is the commonest thing
 * anybody pastes into this add-in — were invisible to the engine. Reported from
 * a real run, on a deck whose slides plainly carried them, where the pane
 * answered "these slides carry no fields yet" about fields it had itself put
 * there.
 *
 * The rule is now stated the other way round, which is what it always meant: a
 * name is anything that is not a brace, a pipe or a line break, and it has to
 * contain **at least one letter or digit**. That keeps the refusals the
 * previous widening was written for — `{{ }}` and `{{!!}}` are not fields —
 * while `{{a b}}`, which was swept in with them, is one. Leading and trailing
 * whitespace is eaten by the `\s*` either side, so `{{ Name }}` and `{{Name}}`
 * are the same field and match a header the parse has already trimmed.
 *
 * Excluding `{`, `}` and `|` is what keeps it from running away: a match cannot
 * cross into the next placeholder, and the format pipe still splits. The one
 * thing it costs is a paragraph writing ABOUT the syntax — "use {{ and }} to
 * mark a field" now reads as a field called "and". That is a template author
 * documenting the tool inside the tool, and it is worth the pivot headers.
 */
export const FIELD = /\{\{\s*([^{}|\r\n]*?[\p{L}\p{N}][^{}|\r\n]*?)\s*(?:\|\s*([^{}\r\n]+?)\s*)?\}\}/gu;

/**
 * Whether a column name can be written as a field at all.
 *
 * The pane offers a button per column that puts `{{Column}}` on the slide, and
 * the engine reads it back with `FIELD`. Those two must agree, and for an hour
 * they did not: the button happily produced a token the reader could not see.
 * A shared function is the only version of "these two agree" that cannot rot —
 * a second predicate in the pane would drift the first time this pattern moved.
 *
 * Answers by ASKING the reader rather than by restating it: the token is built
 * and matched, and the name that comes back has to be the column that went in.
 * That catches the interesting case as well as the obvious one — a header
 * carrying a pipe or a brace does not merely fail to match, it matches a
 * DIFFERENT, shorter name, which would put a field on the slide that silently
 * binds to nothing.
 */
export function canBeField(column: string): boolean {
  const token = `{{${column}}}`;
  // A fresh regex: `FIELD` is global and carries `lastIndex` between calls.
  const hits = [...token.matchAll(new RegExp(FIELD.source, FIELD.flags))];
  return hits.length === 1 && hits[0]?.[0] === token && hits[0]?.[1] === column;
}

/** Answers the value for a field, or null to leave the placeholder visible. */
export type Resolve = (name: string, format?: string) => string | null;

interface Span {
  node: Element;
  from: number;
  to: number;
}

function spansOf(nodes: Element[]): { spans: Span[]; joined: string } {
  const spans: Span[] = [];
  let joined = "";
  for (const node of nodes) {
    const s = node.textContent ?? "";
    spans.push({ node, from: joined.length, to: joined.length + s.length });
    joined += s;
  }
  return { spans, joined };
}

/** The text nodes of a DrawingML paragraph, which is where a slide's placeholders live. */
function runsOf(paragraph: Element): Element[] {
  return elements(paragraph, A_NS, "t");
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
  return mergeRuns(runsOf(paragraph), resolve);
}

/**
 * Merge one run of text held across several nodes. Returns true if anything changed.
 *
 * The engine of `mergeParagraph`, taken out of it because a placeholder does
 * not only live in a DrawingML paragraph. A chart's category labels are
 * `<c:v>` inside a string cache, and the workbook behind that chart keeps the
 * same strings in `<si>` elements that may themselves be split into runs. All
 * three are the same problem — text a merge has to find across nodes it did not
 * choose the boundaries of — and one implementation is the only version of
 * "they behave the same" that cannot rot.
 *
 * The caller says which nodes make up the text; everything else is identical,
 * including the `xml:space` guard, because a chart label ending in a space is
 * the same support ticket as a slide's.
 */
export function mergeRuns(nodes: Element[], resolve: Resolve): boolean {
  const { spans, joined } = spansOf(nodes);
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
  for (const nodes of textGroups(doc)) if (mergeRuns(nodes, resolve)) n++;
  return n;
}

/**
 * Every group of nodes in a part whose text is read as one string.
 *
 * A DrawingML paragraph reaches shape bodies, table cells, grouped shapes,
 * speaker notes, a chart's titles and a SmartArt node's label, because they all
 * hold ordinary `<a:p>` elements at some depth. Two more kinds are here because
 * merging a chart means reaching text that is not in a paragraph at all:
 *
 * - a **chart's cached strings**, the series names and category labels, which
 *   are `<c:v>` inside `<c:strCache>` or `<c:strLit>`. NOT every `<c:v>`: the
 *   same element holds a chart's NUMBERS inside `<c:numCache>`, where the text
 *   has to parse as a number and a placeholder cannot go.
 * - a **workbook's shared strings**, `<si>`, which Excel splits into `<r><t>`
 *   runs exactly as PowerPoint splits a paragraph.
 *
 * One reader for the three, so `fieldsIn` and `mergeDocument` cannot come apart
 * about what a placeholder is — the pane counts what the engine will fill by
 * asking the engine.
 */
function textGroups(doc: Document): Element[][] {
  const out: Element[][] = [];
  for (const paragraph of elements(doc, A_NS, "p")) out.push(runsOf(paragraph));
  for (const cache of [...elements(doc, C_NS, "strCache"), ...elements(doc, C_NS, "strLit")]) {
    for (const v of elements(cache, C_NS, "v")) out.push([v]);
  }
  // A series name written literally rather than referenced: `<c:tx><c:v>`.
  for (const tx of elements(doc, C_NS, "tx")) for (const v of children(tx, C_NS, "v")) out.push([v]);
  for (const si of elements(doc, SSML_NS, "si")) out.push(elements(si, SSML_NS, "t"));
  return out;
}

/** Every field name a part refers to, in first-seen order. For the pane's field list. */
export function fieldsIn(doc: Document): string[] {
  const seen = new Set<string>();
  for (const nodes of textGroups(doc)) {
    const { joined } = spansOf(nodes);
    for (const hit of joined.matchAll(FIELD)) if (hit[1]) seen.add(hit[1]);
  }
  return [...seen];
}
