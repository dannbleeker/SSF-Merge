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

/** Resolve a relationship target against the part that owns the relationship. */
export function resolvePart(ownerPart, target) {
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

/**
 * Every relationship id a part's markup names, with the element that named it.
 *
 * Matched on the `r:` prefix rather than on a resolved namespace so this can
 * read raw text. Every producer writes the relationship namespace as `r:`; a
 * part that bound it to another prefix would be missed, which costs a check and
 * never invents one.
 */
export function referencesIn(xml) {
  const out = [];
  for (const m of xml.matchAll(/<([A-Za-z0-9]+:)?([A-Za-z0-9]+)\b([^>]*)>/g)) {
    const [, prefix = "", element = "", attrs = ""] = m;
    for (const a of attrs.matchAll(/\br:([A-Za-z]+)="([^"]*)"/g)) {
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
      if (!names.has(target)) problems.push(`${name}: ${id} points at ${rel.target}, which is not in the package`);
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
  return problems;
}
