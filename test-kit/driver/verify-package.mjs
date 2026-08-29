#!/usr/bin/env node
/**
 * Verify a merged SSF Merge test-kit deck as a PACKAGE.
 *
 * Usage: node test-kit/driver/verify-package.mjs <deck.pptx>
 *
 * This is deliberately an alignment checker, not a counter. Three charts each
 * holding a distinct region proves nothing about WHICH slide got which region;
 * every per-row check below is anchored to the slide's own title text, so a
 * correctly-shaped deck with the rows shuffled fails.
 *
 * Every check names the condition it tested, not the cause it guessed at.
 * A check that cannot run at all reports N/A and is counted as a failure to
 * observe, never as a pass.
 */
import { readFileSync } from "node:fs";
import { createHash } from "node:crypto";
import JSZip from "jszip";
import { packageProblems } from "../../scripts/package-integrity.mjs";

const deckPath = process.argv[2];
if (!deckPath) {
  console.error("usage: node verify-package.mjs <deck.pptx>");
  process.exit(2);
}

/** The three rows of test-kit/data.txt, in order. */
const ROWS = [
  { name: "Ada", region: "Nordics", revenue: "1 250 000", renewal: "1 Mar 2026", short: "1 Mar", photo: "ada.png" },
  { name: "Grace", region: "Benelux", revenue: "880 000", renewal: "15 Apr 2026", short: "15 Apr", photo: "grace.png" },
  { name: "Alan", region: "DACH", revenue: "1 640 000", renewal: "30 May 2026", short: "30 May", photo: "alan.png" },
];

const results = [];
function record(check, ok, detail) {
  results.push({ check, ok, detail });
}

// ---------------------------------------------------------------- helpers

const REL_RE = /<Relationship\b[^>]*>/g;
const attr = (tag, name) => {
  const m = new RegExp(`\\b${name}="([^"]*)"`).exec(tag);
  return m ? m[1] : null;
};

/** Resolve a relationship Target against a package directory ("" = root). */
function resolveFromDir(dir, target) {
  if (target.startsWith("/")) return target.slice(1);
  const segments = dir === "" ? [] : dir.split("/").filter(Boolean);
  for (const seg of target.split("/")) {
    if (seg === "..") segments.pop();
    else if (seg !== "." && seg !== "") segments.push(seg);
  }
  return segments.join("/");
}

/** Resolve a relationship Target against the directory of the owning part. */
function resolveTarget(ownerPart, target) {
  const i = ownerPart.lastIndexOf("/");
  return resolveFromDir(i === -1 ? "" : ownerPart.slice(0, i), target);
}

function relsPathFor(part) {
  const i = part.lastIndexOf("/");
  return `${part.slice(0, i)}/_rels/${part.slice(i + 1)}.rels`;
}

async function readRels(zip, part) {
  const f = zip.file(relsPathFor(part));
  if (!f) return [];
  const xml = await f.async("string");
  return (xml.match(REL_RE) ?? []).map((tag) => ({
    id: attr(tag, "Id"),
    type: attr(tag, "Type") ?? "",
    target: attr(tag, "Target") ?? "",
    external: (attr(tag, "TargetMode") ?? "") === "External",
    owner: part,
  }));
}

/**
 * All <a:t> and <c:v> text in a part, joined.
 *
 * The tag may carry attributes — `<a:t xml:space="preserve"> 1 Mar</a:t>` is
 * how PowerPoint writes a run with leading space, and a pattern that only
 * matches the bare tag silently drops exactly those runs. That mis-read a
 * correctly-merged date as a lost one once already; do not tighten this back.
 */
const textOf = (xml) =>
  [...xml.matchAll(/<(?:a:t|c:v|t)(?:\s[^>]*)?>([^<]*)<\/(?:a:t|c:v|t)>/g)].map((m) => m[1]).join(" · ");

const sha = (buf) => createHash("sha256").update(buf).digest("hex").slice(0, 12);

// ---------------------------------------------------------------- load

const zip = await JSZip.loadAsync(readFileSync(deckPath));
const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);

const slidePaths = names
  .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
  .sort((a, b) => Number(/(\d+)/.exec(a)[1]) - Number(/(\d+)/.exec(b)[1]));

console.log(`\ndeck        : ${deckPath}`);
console.log(`parts       : ${names.length}`);
console.log(
  `slides      : ${slidePaths.length}  (${slidePaths.map((p) => p.replace("ppt/slides/", "")).join(", ")})\n`,
);

// ------------------------------------------- 1. the package agrees with itself

{
  // One implementation, shared with `test/integrity.test.ts`. This block used
  // to walk the rels itself and check only that every Target resolved — which
  // is one of the three ways a package can contradict itself, and not the one
  // that shipped twice. Markup naming a relationship the part no longer has,
  // and a reference that resolves to the WRONG KIND of part, are the other two.
  const parts = new Map();
  for (const name of names) {
    const file = zip.file(name);
    if (!file) continue;
    const xml = name.endsWith(".xml") || name.endsWith(".rels");
    parts.set(name, xml ? await file.async("string") : await file.async("uint8array"));
  }
  const found = packageProblems(parts);
  record(
    "the package agrees with itself: relationships, references and content types",
    found.length === 0,
    found.length === 0
      ? `${parts.size} parts, no dangling target, no unresolvable reference, no reference to the wrong kind of part`
      : `${found.length} problem(s): ${found.slice(0, 5).join("; ")}`,
  );
}

// ------------------------------------------- 2. part counts

const chartParts = names.filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n)).sort();
const workbookParts = names.filter((n) => /^ppt\/embeddings\/.+\.xlsx$/.test(n)).sort();
const mediaParts = names.filter((n) => /^ppt\/media\//.test(n)).sort();

record("three chart parts", chartParts.length === 3, `${chartParts.length}: ${chartParts.join(", ") || "none"}`);
record(
  "three embedded workbooks",
  workbookParts.length === 3,
  `${workbookParts.length}: ${workbookParts.join(", ") || "none"}`,
);
record("three media parts", mediaParts.length === 3, `${mediaParts.length}: ${mediaParts.join(", ") || "none"}`);

// ------------------------------------------- per-slide alignment

/** Map each merged slide to the row its title text names. */
const slideInfo = [];
for (const sp of slidePaths) {
  const xml = await zip.file(sp).async("string");
  const text = textOf(xml);
  const rels = await readRels(zip, sp);
  const row = ROWS.find((r) => new RegExp(`(^|[^A-Za-z])${r.name}([^A-Za-z]|$)`).test(text)) ?? null;
  slideInfo.push({ path: sp, xml, text, rels, row });
}

const unattributed = slideInfo.filter((s) => !s.row);
record(
  "every slide's text names exactly one of the three rows",
  unattributed.length === 0,
  unattributed.length === 0
    ? slideInfo.map((s) => `${s.path.replace("ppt/slides/", "")}=${s.row.name}`).join(", ")
    : `${unattributed.length} slide(s) name no row: ${unattributed.map((s) => s.path).join(", ")}`,
);

// ------------------------------------------- 3. charts: title + strCache

{
  const detail = [];
  let ok = true;
  let seen = 0;
  for (const s of slideInfo) {
    const chartRels = s.rels.filter((r) => r.type.endsWith("/chart"));
    for (const cr of chartRels) {
      seen++;
      const cp = resolveTarget(s.path, cr.target);
      const cxml = await zip.file(cp)?.async("string");
      if (!cxml) {
        ok = false;
        detail.push(`${s.path} -> ${cp} MISSING`);
        continue;
      }
      const ctext = textOf(cxml);
      const hasOwn = s.row && ctext.includes(s.row.region);
      const hasPlaceholder = ctext.includes("{{Region}}");
      const foreign = ROWS.filter((r) => r !== s.row && ctext.includes(r.region)).map((r) => r.region);
      if (!hasOwn || hasPlaceholder || foreign.length) ok = false;
      detail.push(
        `${s.path.replace("ppt/slides/", "")}->${cp.replace("ppt/charts/", "")} own=${hasOwn ? s.row.region : "NO"}` +
          (hasPlaceholder ? " PLACEHOLDER-LEFT" : "") +
          (foreign.length ? ` FOREIGN=${foreign.join("/")}` : ""),
      );
    }
  }
  if (seen === 0) ok = false;
  record(
    "each chart holds its OWN slide's region, no {{Region}}, no other row's",
    ok,
    seen === 0 ? "no chart relationship found on any slide" : detail.join(" | "),
  );
}

// ------------------------------------------- 4. workbooks: sharedStrings

{
  const detail = [];
  let ok = true;
  let seen = 0;
  for (const s of slideInfo) {
    for (const cr of s.rels.filter((r) => r.type.endsWith("/chart"))) {
      const cp = resolveTarget(s.path, cr.target);
      const crels = await readRels(zip, cp);
      for (const wr of crels.filter((r) => r.type.endsWith("/package") || /\.xlsx$/.test(r.target))) {
        seen++;
        const wp = resolveTarget(cp, wr.target);
        const buf = await zip.file(wp)?.async("nodebuffer");
        if (!buf) {
          ok = false;
          detail.push(`${wp} MISSING`);
          continue;
        }
        let inner;
        try {
          inner = await JSZip.loadAsync(buf);
        } catch (e) {
          ok = false;
          detail.push(`${wp} UNREADABLE (${e.message})`);
          continue;
        }
        const ssFile = inner.file("xl/sharedStrings.xml");
        const sheetFiles = Object.keys(inner.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
        let body = ssFile ? await ssFile.async("string") : "";
        for (const sf of sheetFiles) body += await inner.file(sf).async("string");
        const hasOwn = s.row && body.includes(s.row.region);
        const hasPlaceholder = body.includes("{{Region}}");
        const foreign = ROWS.filter((r) => r !== s.row && body.includes(r.region)).map((r) => r.region);
        if (!hasOwn || hasPlaceholder || foreign.length) ok = false;
        detail.push(
          `${wp.replace("ppt/embeddings/", "")} own=${hasOwn ? s.row.region : "NO"}` +
            (ssFile ? "" : " (no sharedStrings; searched sheets)") +
            (hasPlaceholder ? " PLACEHOLDER-LEFT" : "") +
            (foreign.length ? ` FOREIGN=${foreign.join("/")}` : ""),
        );
      }
    }
  }
  if (seen === 0) ok = false;
  record(
    "each chart's workbook holds that slide's region",
    ok,
    seen === 0 ? "no workbook relationship found from any chart" : detail.join(" | "),
  );
}

// ------------------------------------------- 5. SmartArt drawing part

{
  const detail = [];
  let ok = true;
  let seen = 0;
  for (const s of slideInfo) {
    const drawRels = s.rels.filter((r) => r.type.endsWith("/diagramDrawing"));
    const dataRels = s.rels.filter((r) => r.type.endsWith("/diagramData"));
    if (!dataRels.length) continue;
    seen++;
    if (!drawRels.length) {
      ok = false;
      detail.push(`${s.path.replace("ppt/slides/", "")} has diagramData but NO diagramDrawing`);
      continue;
    }
    for (const dr of drawRels) {
      const dp = resolveTarget(s.path, dr.target);
      const dxml = await zip.file(dp)?.async("string");
      if (!dxml) {
        ok = false;
        detail.push(`${dp} MISSING`);
        continue;
      }
      const dtext = textOf(dxml);
      const hasOwn = s.row && dtext.includes(s.row.name);
      const hasPlaceholder = /\{\{Name\}\}/.test(dtext);
      if (!hasOwn || hasPlaceholder) ok = false;
      detail.push(
        `${dp.replace("ppt/diagrams/", "")} own=${hasOwn ? s.row.name : "NO"}` +
          (hasPlaceholder ? " PLACEHOLDER-LEFT" : ""),
      );
    }
  }
  if (seen === 0) {
    record(
      "SmartArt drawing part holds the row's name",
      false,
      "N/A - no slide carries a diagramData relationship, so there is no SmartArt in this deck to check",
    );
  } else {
    record("SmartArt drawing part (drawingN.xml) holds the row's name", ok, detail.join(" | "));
  }
}

// ------------------------------------------- 6. {{Nickname}} survives

{
  const withNickname = slideInfo.filter((s) => s.text.includes("{{Nickname}}"));
  record(
    "{{Nickname}} still present (no such column, must not blank out)",
    withNickname.length === 3,
    `${withNickname.length} slide(s) carry it: ${withNickname.map((s) => s.path.replace("ppt/slides/", "")).join(", ") || "none"}`,
  );
}

// ------------------------------------------- 7. notes pages

{
  const detail = [];
  let ok = true;
  let seen = 0;
  for (const s of slideInfo) {
    // A deck that still carries its template slides has slides that name no
    // row. They are not failures; they simply are not this check's business.
    if (!s.row) continue;
    for (const nr of s.rels.filter((r) => r.type.endsWith("/notesSlide"))) {
      seen++;
      const np = resolveTarget(s.path, nr.target);
      const nxml = await zip.file(np)?.async("string");
      if (!nxml) {
        ok = false;
        detail.push(`${np} MISSING`);
        continue;
      }
      const ntext = textOf(nxml).replace(/\s*·\s*/g, "");
      const want = `Call ${s.row.name} before ${s.row.short}`;
      const hit = ntext.includes(want);
      if (!hit) ok = false;
      detail.push(
        `${np.replace("ppt/notesSlides/", "")} ${hit ? "OK" : "WANT<" + want + "> GOT<" + ntext.slice(0, 60) + ">"}`,
      );
    }
  }
  if (seen === 0) ok = false;
  record(
    'notes pages read "Call <Name> before <date>"',
    ok,
    seen === 0 ? "no notesSlide relationship on any slide" : detail.join(" | "),
  );
}

// ------------------------------------------- 8. formats on the slide text

{
  const detail = [];
  let ok = true;
  for (const s of slideInfo) {
    if (!s.row) continue;
    const flat = s.text.replace(/\s*·\s*/g, "");
    const wantsRevenue = flat.includes("EUR");
    if (!wantsRevenue) continue;
    const rev = flat.includes(s.row.revenue.replace(/ /g, " ")) || flat.includes(s.row.revenue);
    const ren = flat.includes(s.row.renewal);
    if (!rev || !ren) ok = false;
    detail.push(
      `${s.path.replace("ppt/slides/", "")} revenue=${rev ? "OK" : "NO(" + s.row.revenue + ")"} renewal=${ren ? "OK" : "NO(" + s.row.renewal + ")"}`,
    );
  }
  record(
    "number and date formats rendered",
    ok && detail.length === 3,
    detail.join(" | ") || "no slide carried a revenue line",
  );
}

// ------------------------------------------- 9. pictures: each slide its own

{
  const detail = [];
  let ok = true;
  const seenHashes = new Map();
  for (const s of slideInfo) {
    for (const ir of s.rels.filter((r) => r.type.endsWith("/image") && !r.external)) {
      const ip = resolveTarget(s.path, ir.target);
      const buf = await zip.file(ip)?.async("nodebuffer");
      if (!buf) {
        ok = false;
        detail.push(`${ip} MISSING`);
        continue;
      }
      const h = sha(buf);
      seenHashes.set(h, (seenHashes.get(h) ?? 0) + 1);
      detail.push(
        `${s.path.replace("ppt/slides/", "")}(${s.row?.name}) -> ${ip.replace("ppt/media/", "")} sha=${h} ${buf.length}B`,
      );
    }
  }
  const distinct = seenHashes.size;
  if (distinct !== 3) ok = false;
  record("three distinct image payloads, one per row", ok, `${distinct} distinct sha(s). ${detail.join(" | ")}`);
}

// ------------------------------------------- source PNG identity

{
  const want = new Map();
  for (const r of ROWS) {
    try {
      want.set(sha(readFileSync(`test-kit/${r.photo}`)), r);
    } catch {
      /* source not readable from here */
    }
  }
  if (want.size === 3) {
    const detail = [];
    let ok = true;
    for (const s of slideInfo) {
      for (const ir of s.rels.filter((r) => r.type.endsWith("/image") && !r.external)) {
        const ip = resolveTarget(s.path, ir.target);
        const h = sha(await zip.file(ip).async("nodebuffer"));
        const src = want.get(h);
        const right = src && s.row && src.name === s.row.name;
        if (!right) ok = false;
        detail.push(`${s.row?.name}->${src ? src.photo : "UNRECOGNISED-BYTES"}${right ? "" : " MISMATCH"}`);
      }
    }
    record("each slide carries ITS OWN row's source PNG, byte-identical", ok, detail.join(" | "));
  } else {
    record(
      "each slide carries ITS OWN row's source PNG, byte-identical",
      false,
      "N/A - could not read the three source PNGs from test-kit/",
    );
  }
}

// ---------------------------------------------------------------- table

const width = Math.max(...results.map((r) => r.check.length));
console.log("=".repeat(width + 10));
for (const r of results) {
  console.log(`${r.ok ? "PASS" : "FAIL"}  ${r.check.padEnd(width)}`);
  console.log(`      ${r.detail}\n`);
}
const failed = results.filter((r) => !r.ok);
console.log("=".repeat(width + 10));
console.log(`${results.length - failed.length}/${results.length} checks passed.`);
if (failed.length) console.log(`FAILED: ${failed.map((r) => r.check).join(" ; ")}`);
process.exit(failed.length ? 1 : 0);
