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
import { A_NS, C_NS, CX_NS, SSML_NS, children, elements, xmlSafe } from "../pptx/xml.js";

/**
 * One `{{Field}}` or `{{Field|format}}`, found in a piece of text.
 *
 * `index` and `length` are offsets into the text that was scanned, which is the
 * paragraph's JOINED text rather than any one run — see the note at the top of
 * this file for why the two are never the same thing.
 */
export interface FieldHit {
  /** The field name, trimmed. Matches a column header as the parse trimmed it. */
  name: string;
  /** What follows the `|`, trimmed. Absent when the placeholder carries no format. */
  format?: string;
  /** Offset of the opening `{{`. */
  index: number;
  /** Length of the whole placeholder, braces included. */
  length: number;
}

/** A field name must carry at least one of these, or `{{ }}` and `{{!!}}` would be fields. */
const ALNUM = /[\p{L}\p{N}]/u;

/** Neither a name nor a format may contain one: a placeholder lives on one line. */
const BREAK = /[\r\n]/;

/**
 * Every placeholder in a piece of text, in order.
 *
 * **A name is anything that is not a brace, a pipe or a line break, and it has
 * to contain at least one letter or digit.** Unicode, not ASCII: `[\w.]` matched
 * none of `ø`, `å`, `é`, `ü`, Greek or Cyrillic, so `{{Beløb}}` and `{{Måned}}`
 * were INVISIBLE to the engine — never reported to the pane, so never flagged as
 * unmatched either, and the literal braces printed on every merged slide. On a
 * product whose first users write Danish, that is most of a template.
 *
 * **Spaces are part of a name**, which the rule refused for a while by listing
 * the characters it allowed. `{{Row Labels}}`, `{{Min. of cost}}` and
 * `{{Sum of quantity monthly}}` — the literal default headers of an Excel pivot
 * table, which is the commonest thing anybody pastes into this add-in — were
 * invisible for the same reason, reported from a real run on a deck whose slides
 * plainly carried them, where the pane answered "these slides carry no fields
 * yet" about fields it had itself put there. Stating the rule as a refusal keeps
 * `{{ }}` and `{{!!}}` out while letting `{{a b}}` in.
 *
 * Whitespace at either end belongs to nobody: `{{ Name }}` and `{{Name}}` are
 * the same field, and both match a header the parse has already trimmed.
 *
 * Excluding braces and the pipe is what keeps a placeholder from running away: a
 * hit cannot cross into the next one, and the format pipe still splits. The one
 * thing it costs is a paragraph writing ABOUT the syntax — "use {{ and }} to
 * mark a field" reads as a field called "and". That is a template author
 * documenting the tool inside the tool, and it is worth the pivot headers.
 *
 * **A SCAN rather than a regular expression, and that is the point.** The
 * pattern this replaces was
 *
 *     /\{\{\s*([^{}|\r\n]*?[\p{L}\p{N}][^{}|\r\n]*?)\s*(?:\|\s*([^{}\r\n]+?)\s*)?\}\}/gu
 *
 * — two lazy unbounded classes with a required character between them, which is
 * the classic shape for catastrophic backtracking. On a placeholder that opens
 * and never closes, the engine tries every split of the run between the two:
 * `{{` and forty thousand letters took 4.4 seconds, and `{{a|` and four thousand
 * spaces took THIRTY-THREE, synchronously, on text that arrives inside a deck
 * somebody was sent. Three replacement patterns were written and measured for
 * exact equivalence against a corpus, and each fixed one shape and made another
 * worse; the ambiguity is not an accident of how the pattern was written, it IS
 * the trimming behaviour, and a backtracking engine cannot have one without the
 * other. Scanning is O(n) by construction — the same two inputs are 0.4 ms and
 * 0.3 ms — and it is shorter to read than the pattern was.
 *
 * The one deliberate difference in what it ANSWERS is a format made of nothing
 * but whitespace: `{{Total|  }}` reported a format of `" "` and now reports
 * none. Which spans are fields is unchanged, including the corners nobody would
 * design — `{{a|}}` and `{{a|\n}}` are still not placeholders — and
 * `test/text.test.ts` holds both halves against the old pattern over a corpus.
 * The product cannot see the difference: `applyFormat` and `imageMode` both
 * answer a whitespace format exactly as they answer no format, which is asserted
 * rather than assumed.
 */
export function fieldsInText(text: string): FieldHit[] {
  const out: FieldHit[] = [];
  let at = 0;
  for (;;) {
    const open = text.indexOf("{{", at);
    if (open < 0) return out;
    const hit = readField(text, open);
    // Not `open + 2`: the text that defeated this placeholder may itself open
    // one. `{{a{{b}}` has no field called `a` and does have one called `b`.
    if (!hit) {
      at = open + 1;
      continue;
    }
    out.push(hit);
    at = hit.index + hit.length;
  }
}

/**
 * The placeholder opening at `open`, or nothing if that `{{` does not start one.
 *
 * Linear in the distance to the first character that cannot be inside a
 * placeholder, and every one of those characters is either the end of this
 * search or the start of the next one — so the whole scan is linear in the text
 * however many `{{` it holds.
 */
function readField(text: string, open: number): FieldHit | undefined {
  let close = -1;
  for (let i = open + 2; i < text.length; i++) {
    const c = text[i];
    // A brace inside a placeholder ends it, and a lone `}` is not an end:
    // `{{a}b}}` is not a field, which is what the excluded characters said.
    if (c === "{") return undefined;
    if (c === "}") {
      if (text[i + 1] !== "}") return undefined;
      close = i;
      break;
    }
  }
  if (close < 0) return undefined;

  const body = text.slice(open + 2, close);
  const bar = body.indexOf("|");
  const name = (bar < 0 ? body : body.slice(0, bar)).trim();
  if (!ALNUM.test(name) || BREAK.test(name)) return undefined;

  let format: string | undefined;
  if (bar >= 0) {
    const rest = body.slice(bar + 1);
    // A pipe with nothing usable after it is not a placeholder at all — the
    // format is required once the pipe is written. Inherited exactly: it is why
    // `{{Name|}}` prints its own braces instead of merging.
    if (!/[^\r\n]/.test(rest)) return undefined;
    const trimmed = rest.trim();
    if (BREAK.test(trimmed)) return undefined;
    if (trimmed !== "") format = trimmed;
  }
  return { name, format, index: open, length: close + 2 - open };
}

/**
 * Whether a column name can be written as a field at all.
 *
 * The pane offers a button per column that puts `{{Column}}` on the slide, and
 * the engine reads it back with the scanner above. Those two must agree, and for
 * an hour they did not: the button happily produced a token the reader could not
 * see. A shared function is the only version of "these two agree" that cannot
 * rot — a second predicate in the pane would drift the first time this moved.
 *
 * Answers by ASKING the reader rather than by restating it: the token is built
 * and scanned, and the name that comes back has to be the column that went in,
 * having taken the whole token. That catches the interesting case as well as the
 * obvious one — a header carrying a pipe or a brace does not merely fail to
 * match, it matches a DIFFERENT, shorter name, which would put a field on the
 * slide that silently binds to nothing.
 */
export function canBeField(column: string): boolean {
  return whyNotAField(column) === null;
}

/**
 * WHICH rule a column name breaks, as a clause the pane can show.
 *
 * `canBeField` answers yes or no, and the pane printed one sentence for every
 * no: "a field name may not contain a brace or a pipe". Three of the six
 * refusals this file's own tests list are nothing of the kind — `""`, `"   "`
 * and `"!!"` are refused for having no letter or digit in them — and the
 * commonest one in real data is neither.
 *
 * A header cell holding Alt+Enter is ordinary in a spreadsheet — `Revenue`
 * over `(EUR)` — and the clipboard carries it quoted, so the paste parses
 * perfectly and the column is called `Revenue\n(EUR)`. Every row of it is
 * there; the pane refused to offer a chip for it and told the user to go and
 * look for a brace or a pipe in a header that has neither. That is the failure
 * `prepareBlock` states the standard against: every refusal has to be a thing
 * the user can act on.
 *
 * The ASK comes first and decides — the token is built and scanned, so the rule
 * is still the engine's own reader rather than a second opinion. What follows
 * only explains an answer already given, which is why a clause it cannot
 * account for says so instead of naming a rule it has not checked.
 */
export function whyNotAField(column: string): string | null {
  const token = `{{${column}}}`;
  const hits = fieldsInText(token);
  const whole =
    hits.length === 1 && hits[0]?.index === 0 && hits[0]?.length === token.length && hits[0]?.name === column;
  if (whole) return null;
  if (/[{}|]/.test(column)) return "a field name may not contain a brace or a pipe";
  if (BREAK.test(column)) return "a field name lives on one line, and this header has a line break in it";
  if (!ALNUM.test(column.trim())) return "a field name needs at least one letter or digit in it";
  if (column.trim() !== column) return "a field name may not start or end with a space";
  // Nothing reaches this today — the four rules above are the whole of what
  // `readField` refuses, and `toRecordSet` has already trimmed every header.
  // It says what it knows rather than naming a rule it has not checked, which
  // is the mistake the single hardcoded sentence made.
  return "the reader cannot take this name";
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

  const edits: Edit[] = [];
  for (const hit of fieldsInText(joined)) {
    const value = resolve(hit.name, hit.format);
    // A field nobody can resolve stays visible. Blanking it hides the author's
    // typo behind 240 slides that look finished and are not.
    if (value === null) continue;
    edits.push({ start: hit.index, end: hit.index + hit.length, value });
  }
  return editRuns(nodes, edits);
}

/** A replacement, given in offsets into the JOINED text rather than into any one node. */
export interface Edit {
  start: number;
  /** Exclusive. */
  end: number;
  value: string;
}

/**
 * Apply edits to text held across several nodes. Returns true if any were.
 *
 * Taken out of `mergeRuns` because the image pass needs exactly this and had
 * been doing something coarser: it blanked EVERY text node in the paragraph
 * once it had placed a picture. A caption beside the placeholder, or a second
 * field, went with it — silently, on a slide that then looked finished.
 *
 * The replacement lands in the FIRST node the edit touches and the covered
 * characters are removed from the rest, so a value inherits the run properties
 * of the run its placeholder started in. That is the contract a template author
 * works against: format the opening brace the way you want the value to look.
 *
 * Per-character buffers, so every offset keeps referring to the ORIGINAL joined
 * text however many edits one paragraph takes.
 */
export function editRuns(nodes: Element[], edits: Edit[]): boolean {
  if (!edits.length) return false;
  const { spans } = spansOf(nodes);
  if (!spans.length) return false;

  const bufs = spans.map((s) => (s.node.textContent ?? "").split(""));
  for (const edit of edits) {
    let carried = false;
    spans.forEach((s, i) => {
      if (s.to <= edit.start || s.from >= edit.end) return;
      const buf = bufs[i];
      if (!buf) return;
      const lo = Math.max(edit.start, s.from) - s.from;
      const hi = Math.min(edit.end, s.to) - s.from;
      for (let k = lo; k < hi; k++) buf[k] = "";
      if (!carried) {
        buf[lo] = edit.value;
        carried = true;
      }
    });
  }

  spans.forEach((s, i) => {
    const text = xmlSafe((bufs[i] ?? []).join(""));
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
  // A cell that holds its string INLINE instead of pointing at that table:
  // `<c t="inlineStr"><is><t>`. Same `<r><t>` runs, one letter apart, and it
  // was missing — so `mergeWorkbook` opened every worksheet of every embedded
  // workbook, walked it, and could never find anything. Its own comment said
  // the worksheets were read "because a cell may hold its string INLINE", which
  // is the population a generator that never built a shared-string table
  // produces, and that is exactly what nothing here could reach: a label cell
  // written inline came out of the merge still reading `{{Name}}`, in the
  // workbook Excel opens on Edit Data, with the deck itself looking finished.
  for (const is of elements(doc, SSML_NS, "is")) out.push(elements(is, SSML_NS, "t"));
  // A MODERN chart, which keeps its text in two more places of its own.
  //
  // `<cx:pt>` inside a `<cx:strDim>` is a category label. Scoped by the DIM and
  // never by the element name, because `<cx:numDim>` holds the values the chart
  // PLOTS in `<cx:pt>` too — filling one of those with "Nordics" produces a
  // chart PowerPoint reads as corrupt data. Exactly the `<c:v>` distinction two
  // loops above, one namespace over.
  for (const dim of elements(doc, CX_NS, "strDim")) {
    for (const pt of elements(dim, CX_NS, "pt")) out.push([pt]);
  }
  // A series name, and sometimes a title: `<cx:tx><cx:txData><cx:v>`. Sometimes
  // is the operative word — the same title is DrawingML inside `<cx:txPr>` in
  // some files, a copy of this value in others, and the `<a:p>` loop at the top
  // covers that half. Both are filled, because which one a given file uses is
  // not knowable from the schema.
  for (const data of elements(doc, CX_NS, "txData")) {
    for (const v of children(data, CX_NS, "v")) out.push([v]);
  }
  return out;
}

/** Every field name a part refers to, in first-seen order. For the pane's field list. */
export function fieldsIn(doc: Document): string[] {
  const seen = new Set<string>();
  for (const nodes of textGroups(doc)) {
    const { joined } = spansOf(nodes);
    for (const hit of fieldsInText(joined)) seen.add(hit.name);
  }
  return [...seen];
}
