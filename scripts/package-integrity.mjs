#!/usr/bin/env node
/**
 * Whether a `.pptx` is INTERNALLY CONSISTENT — a structural check, not a
 * content one.
 *
 * `test-kit/driver/verify-package.mjs` asks whether a merged deck says the
 * right things: this row's region in this row's chart. It knows the test kit's
 * data by heart and cannot be pointed at anything else. This asks a question no
 * deck is exempt from — do the parts, the relationships and the markup agree —
 * and it needs to know nothing about what the deck is for.
 *
 * It exists because two defects found on 2026-08-29 were the same mistake in
 * two places, and neither was reachable by any gate this repository had. A pass
 * that deletes relationships knew ONE of the ways a slide names them, so it
 * deleted the others: a merged copy came out naming a relationship that was
 * gone, and — because deleting a relationship frees its id for the next thing
 * that needs one — another came out naming a relationship that led somewhere
 * else entirely. PowerPoint calls the first a damaged file. The second it opens
 * without complaint.
 *
 * The checks, and what each is for:
 *
 * - **dangling relationship** — a Target naming a part that is not in the
 *   package. The classic damaged-file cause.
 * - **unresolvable reference** — markup naming a relationship id the part does
 *   not have. The same damage seen from the other side, and the side that
 *   catches a relationship deleted out from under live markup.
 * - **mistyped reference** — a reference that resolves to a relationship of the
 *   wrong KIND: `<a:blip r:embed>` pointing at something that is not an image,
 *   `<p:tags r:id>` at something that is not a tag part. This is the one that
 *   catches an id reused after a delete, which the two checks above both pass.
 * - **undeclared part** — a part no content type covers, and an Override naming
 *   a part that is not there. PowerPoint refuses a package with either.
 * - **a character XML cannot carry** — a C0 control, an unpaired surrogate or
 *   `FFFE`/`FFFF` in the markup. The one problem that makes every check above
 *   moot, because a conforming parser refuses the part before it can disagree
 *   about anything in it.
 *
 * Deliberately NOT checked: whether a part is reachable from the presentation.
 * A stranded part is weight rather than damage, `orphanedParts` is where that
 * decision lives, and a checker that called it an error would fail on every
 * legitimately shared part it did not understand.
 */

/** `<a:blip r:embed>` and friends: which element expects which relationship. */
const EXPECTED_TYPE = [
  { element: "blip", attrs: ["embed", "link"], type: "/image" },
  { element: "svgBlip", attrs: ["embed"], type: "/image" },
  { element: "tags", attrs: ["id"], type: "/tags" },
  { element: "chart", attrs: ["id"], type: null }, // c:chart and cx:chart differ; resolved below
  { element: "sldId", attrs: ["id"], type: "/slide" },
  { element: "sldLayoutId", attrs: ["id"], type: "/slideLayout" },
  { element: "sldMasterId", attrs: ["id"], type: "/slideMaster" },
  { element: "notesMasterId", attrs: ["id"], type: "/notesMaster" },
  { element: "relIds", attrs: ["dm"], type: "/diagramData" },
  { element: "relIds", attrs: ["lo"], type: "/diagramLayout" },
  { element: "relIds", attrs: ["qs"], type: "/diagramQuickStyle" },
  { element: "relIds", attrs: ["cs"], type: "/diagramColors" },
];

/**
 * The type a `<chart>` reference wants, which depends on its namespace.
 *
 * `<c:chart>` in a slide's graphicData is the classic chart; `<cx:chart>` is a
 * modern one, under a Microsoft relationship. Same local name, two answers —
 * the overloaded-name trap this project has now met three times.
 */
function chartTypeFor(prefix) {
  return prefix === "cx" ? "/chartEx" : "/chart";
}

/**
 * Whether the package holds this part, under either spelling of its name.
 *
 * A relationship `Target` is a URI reference, so a part name containing a space
 * is written `my%20photo.png` while the zip entry is `my photo.png`. Compared
 * raw, that part is "not in the package" on a package that holds it — the third
 * invented problem of this family, after the empty `r:blip` and the
 * root-relative target, and the same population each time.
 *
 * Decoded AND raw, rather than decoded only. A producer that escaped nothing
 * can leave a literal `%` in a part name, and decoding that one turns a name the
 * package really has into one it does not — trading this false alarm for its
 * mirror image. Accepting either spelling is the only direction that invents
 * nothing.
 *
 * `decodeURIComponent` throws on a lone `%`, which is what an unescaped name
 * looks like, so the throw is the ordinary case here rather than the odd one.
 *
 * The decoding is deliberately NOT in `resolvePart`: that function is pure and
 * is held against the engine's own resolver by a corpus, and whether the ENGINE
 * should decode is a question about every deck a merge touches. It is not
 * answered here.
 *
 * @param {Set<string>} names
 * @param {string} part
 * @returns {boolean}
 */
export function holds(names, part) {
  if (names.has(part)) return true;
  if (!part.includes("%")) return false;
  try {
    return names.has(decodeURIComponent(part));
  } catch {
    return false;
  }
}

/**
 * Resolve a relationship target against the part that owns the relationship.
 *
 * A leading slash means the target is already a part name, given from the
 * package root — legal OOXML, and it is the spelling this checker got wrong.
 * Resolved relatively it became `ppt/slides/ppt/media/image1.png`, a part no
 * package holds, and a sound deck was reported as pointing at something that
 * is not in it. Same shape as the empty `r:blip` of 2026-08-30: a false alarm
 * from a reader, on markup this engine never wrote and a sender's deck may.
 *
 * `src/core/pptx/pkg.ts` has always handled it, and that divergence is the
 * point — two implementations of one rule, one of them wrong for a year.
 * `test/integrity.test.ts` now runs both over the same pairs.
 *
 * @param {string} ownerPart
 * @param {string} target
 * @returns {string}
 */
export function resolvePart(ownerPart, target) {
  if (target.startsWith("/")) return target.slice(1);
  /** @type {string[]} */
  const segments = ownerPart.split("/").slice(0, -1);
  for (const seg of target.split("/")) {
    if (seg === "..") segments.pop();
    else if (seg !== "." && seg !== "") segments.push(seg);
  }
  return segments.join("/");
}

/** `ppt/slides/slide1.xml` → `ppt/slides/_rels/slide1.xml.rels`. */
export function relsPathFor(part) {
  const cut = part.lastIndexOf("/");
  return cut < 0 ? `_rels/${part}.rels` : `${part.slice(0, cut)}/_rels/${part.slice(cut + 1)}.rels`;
}

/**
 * Every relationship a part declares, as `id -> { type, target, external }`.
 *
 * Parsed with a regular expression rather than an XML reader, because this file
 * is imported by a Node script with no DOM and by the suite. A `.rels` part is
 * a flat list of self-closing elements and nothing else, which is the one shape
 * where that is a fair trade.
 */
export function relationshipsOf(relsXml) {
  const out = new Map();
  if (!relsXml) return out;
  for (const m of relsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const tag = m[0];
    const id = /\bId="([^"]*)"/.exec(tag)?.[1];
    if (!id) continue;
    out.set(id, {
      type: /\bType="([^"]*)"/.exec(tag)?.[1] ?? "",
      target: /\bTarget="([^"]*)"/.exec(tag)?.[1] ?? "",
      external: /\bTargetMode="External"/.test(tag),
    });
  }
  return out;
}

/** The relationships namespace, as a part's markup binds it to a prefix. */
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/**
 * Which prefixes THIS part binds to the relationships namespace.
 *
 * Falls back to `r` when the part declares none, which is every fragment and
 * every part that inherits the binding from a parent it is not shown with.
 *
 * @param {string} xml
 * @returns {string[]}
 */
export function relPrefixesIn(xml) {
  const found = [];
  let rBoundElsewhere = false;
  for (const m of xml.matchAll(/xmlns:([A-Za-z0-9_-]+)="([^"]*)"/g)) {
    if (m[2] === REL_NS) {
      if (!found.includes(m[1])) found.push(m[1]);
    } else if (m[1] === "r") rBoundElsewhere = true;
  }
  if (found.length) return found;
  // The fallback is for a part that declares nothing — a fragment, or one that
  // inherits the binding from a parent it is not shown with. It must NOT apply
  // to a part that binds `r` to some other namespace: there `r:embed` is a
  // different attribute that happens to share a name, and reading it as a
  // relationship is how a sound part gets reported as naming one it lacks.
  return rBoundElsewhere ? [] : ["r"];
}

/**
 * Every relationship id a part's markup names, with the element that named it.
 *
 * Read as raw text rather than through a DOM, for the reason `relationshipsIn`
 * gives — but the PREFIX is resolved from the part's own `xmlns` declarations
 * rather than assumed to be `r`.
 *
 * This docstring used to say the prefix was matched literally, that a part
 * binding the namespace to another prefix "would be missed, which costs a check
 * and never invents one". The first half was true and the second was not, in
 * both directions. A part declaring `xmlns:rel="…/relationships"` had every one
 * of its references skipped, so a genuinely dangling one went unreported. And a
 * part binding `r:` to something else entirely — legal XML, and the prefix is
 * only a name — had its attributes read as relationship references and was
 * reported as naming relationships it does not have. That is an invented
 * problem on sound markup, which is the failure this checker can least afford:
 * it is the third of this shape after the empty `r:blip` and the root-relative
 * target, and the same population every time, a deck written by something other
 * than PowerPoint.
 *
 * @param {string} xml
 */
export function referencesIn(xml) {
  /** @type {{ element: string, prefix: string, attr: string, id: string }[]} */
  const out = [];
  const prefixes = relPrefixesIn(xml).map((p) => p.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  // No prefix is bound to the namespace, so the part names no relationship.
  // Building the alternation from an empty list would produce `(?:):(...)`,
  // which matches ANY prefixed attribute — the invented-problem direction,
  // reintroduced by the fix for it.
  if (prefixes.length === 0) return out;
  const attrPattern = new RegExp(`\\b(?:${prefixes.join("|")}):([A-Za-z]+)="([^"]*)"`, "g");
  for (const m of xml.matchAll(/<([A-Za-z0-9]+:)?([A-Za-z0-9]+)\b([^>]*)>/g)) {
    const [, prefix = "", element = "", attrs = ""] = m;
    for (const a of attrs.matchAll(attrPattern)) {
      const id = a[2] ?? "";
      // An EMPTY id names no relationship, so it is not a reference to resolve.
      //
      // `r:blip=""` is the ordinary OOXML idiom for "no image here", and it is
      // PowerPoint's own markup: a SmartArt layout part carries one on every
      // `<dgm:shape>` that has no picture, and that part correctly has no
      // `.rels` beside it at all. Counting them as references asked the
      // relationship map for `""`, got nothing back, and reported a sound
      // package as naming relationships it does not have — four per layout
      // part, on markup this engine never wrote. The real-host round of
      // 2026-08-30 saw sixteen of them and no genuine problem, on a deck
      // PowerPoint then opened with no repair prompt.
      //
      // `relationshipIdsIn` in src/core/pptx/xml.ts collects the same empty
      // value, and is INERT there rather than wrong: that set is asked
      // `has(id)` for each relationship the part really has, and no
      // relationship has an empty Id. It reads markup to decide what to KEEP;
      // this reads it to decide what must RESOLVE, and only the second
      // direction can be misled by an id that was never meant to name anything.
      if (id === "") continue;
      out.push({ element, prefix: prefix.replace(":", ""), attr: a[1] ?? "", id });
    }
  }
  return out;
}

/** What type, if any, a reference is required to lead to. */
export function expectedTypeOf(ref) {
  if (ref.element === "chart" && ref.attr === "id") return chartTypeFor(ref.prefix);
  for (const rule of EXPECTED_TYPE) {
    if (rule.element !== ref.element || !rule.attrs.includes(ref.attr)) continue;
    return rule.type;
  }
  return undefined;
}

/**
 * Check one package.
 *
 * `parts` is a Map of part name to its bytes-or-text; only XML parts are read.
 * Answers a list of problems, each a sentence naming the part and the thing
 * that is wrong with it. An empty list is a package whose internal references
 * all agree.
 */
/**
 * The XML 1.0 `Char` production, as a refusal.
 *
 * Most of the C0 controls, an unpaired surrogate and `FFFE`/`FFFF` may not
 * appear in an XML document at all — not raw, and not as a numeric entity
 * either: `&#11;` is exactly as ill-formed as the character. A part carrying
 * one is a part a conforming parser rejects, and PowerPoint condemns the whole
 * file for it.
 *
 * Worth a package-level check rather than only a guard at the writer, because
 * this repository parses and serialises with `@xmldom/xmldom` at BOTH ends:
 * it writes such a character straight through and reads it straight back, so a
 * round trip proves nothing and every other gate here stays green. The one
 * question that separates the two is whether the bytes themselves are legal,
 * and that is asked of the markup rather than of a document.
 *
 * A well-formed surrogate PAIR is one code point above `FFFF` and is perfectly
 * legal. Under the `u` flag the pattern walks code points, so a pair is never
 * one of these and only an unpaired half matches — which is what a check that
 * fired on every emoji in the deck would have got wrong.
 */
// The rule is aimed at a control character reaching a pattern by accident. Here
// they ARE the subject: this is the set XML refuses, and it cannot be written
// without naming them.
// eslint-disable-next-line no-control-regex
const XML_FORBIDDEN = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|\p{Surrogate}/u;

export function packageProblems(parts) {
  const problems = [];
  const names = new Set(parts.keys());
  const text = (name) => {
    const v = parts.get(name);
    return typeof v === "string" ? v : undefined;
  };

  for (const name of names) {
    if (!name.endsWith(".rels")) continue;
    const owner = name.includes("/_rels/")
      ? `${name.slice(0, name.indexOf("/_rels/"))}/${name.slice(name.indexOf("/_rels/") + 7, -".rels".length)}`
      : name.slice("_rels/".length, -".rels".length);
    for (const [id, rel] of relationshipsOf(text(name) ?? "")) {
      if (rel.external || /^[a-z][a-z0-9+.-]*:/i.test(rel.target)) continue;
      const target = resolvePart(owner, rel.target);
      if (!holds(names, target)) problems.push(`${name}: ${id} points at ${rel.target}, which is not in the package`);
    }
  }

  for (const name of names) {
    if (name.endsWith(".rels") || !name.endsWith(".xml")) continue;
    const body = text(name);
    if (body === undefined) continue;
    const rels = relationshipsOf(text(relsPathFor(name)) ?? "");
    for (const ref of referencesIn(body)) {
      const rel = rels.get(ref.id);
      if (!rel) {
        problems.push(
          `${name}: <${ref.prefix ? `${ref.prefix}:` : ""}${ref.element} r:${ref.attr}="${ref.id}"> names a relationship the part does not have`,
        );
        continue;
      }
      const want = expectedTypeOf(ref);
      if (want && !rel.type.endsWith(want)) {
        problems.push(
          `${name}: <${ref.prefix ? `${ref.prefix}:` : ""}${ref.element} r:${ref.attr}="${ref.id}"> wants ${want.slice(1)} and leads to ${rel.type.split("/").pop()} (${rel.target})`,
        );
      }
    }
  }

  const types = text("[Content_Types].xml") ?? "";
  const defaults = new Set([...types.matchAll(/Default Extension="([^"]*)"/g)].map((m) => (m[1] ?? "").toLowerCase()));
  const overrides = new Set([...types.matchAll(/PartName="\/([^"]*)"/g)].map((m) => m[1] ?? ""));
  for (const part of overrides) {
    if (!names.has(part)) problems.push(`[Content_Types].xml: declares /${part}, which is not in the package`);
  }
  for (const name of names) {
    if (name === "[Content_Types].xml" || overrides.has(name)) continue;
    const ext = (name.split(".").pop() ?? "").toLowerCase();
    if (!defaults.has(ext)) problems.push(`${name}: no content type covers it`);
  }
  // A part whose BYTES XML cannot carry. Last, because it is the one problem
  // that makes every check above meaningless: a parser refuses the part before
  // it can disagree about a relationship.
  for (const name of names) {
    if (!name.endsWith(".xml") && !name.endsWith(".rels")) continue;
    const body = text(name);
    if (body === undefined) continue;
    const at = body.search(XML_FORBIDDEN);
    if (at < 0) continue;
    const code = body.charCodeAt(at).toString(16).padStart(4, "0");
    problems.push(`${name}: holds U+${code.toUpperCase()} at offset ${at}, which XML cannot carry`);
  }
  return problems;
}
