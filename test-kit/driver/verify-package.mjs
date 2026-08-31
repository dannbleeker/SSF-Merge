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

/** Print the table and leave with the right code. Never returns. */
function report() {
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

// The part COUNTS used to be taken here, over the whole package, and asked for
// exactly three of each. That is only true of a deck the template has been swept
// out of: a round done the way the manual asks keeps the template, whose own
// chart, workbook and diagram are parts too — so a correct deck counted four and
// was marked down for it three times. They are counted below instead, over the
// parts the MERGED slides actually reach, which is what "three" was always
// about.

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

/**
 * Which slides this run PRODUCED, as opposed to the block it copied from.
 *
 * A round done the way `docs/TEST-KIT.md` asks leaves the template in the deck:
 * the merge appends, it does not consume. Every per-row check below is about
 * the copies, and the template slides are not failures of them — they are not
 * their business. The notes check already said exactly that and skipped them;
 * the others did not, so the real-host round of 2026-08-30 scored 5/13 on a
 * deck that was entirely correct and had to re-derive every claim by hand.
 *
 * A template slide is one still holding a placeholder for a field that HAS a
 * column. `{{Nickname}}` is deliberately excluded: it has no column, it is MEANT
 * to survive, and every merged copy carries it — keying on it would call the
 * whole deck a template and skip everything.
 */
const TEMPLATE_MARK = /\{\{(?!Nickname\}\})[A-Za-z]/;
/** The same rule, reading back WHICH placeholders, for a report. */
const TEMPLATE_MARKS = /\{\{(?!Nickname\}\})[A-Za-z][^}]*\}\}/g;
for (const s of slideInfo) s.template = TEMPLATE_MARK.test(s.text);

const short = (s) => s.path.replace("ppt/slides/", "");
const mergedInfo = slideInfo.filter((s) => s.row && !s.template);
const templateInfo = slideInfo.filter((s) => !s.row || s.template);

console.log(
  `merged      : ${mergedInfo.length}  (${mergedInfo.map(short).join(", ") || "none"})\n` +
    `template    : ${templateInfo.length}  (${templateInfo.map(short).join(", ") || "none"})` +
    `  — skipped by the per-row checks\n`,
);

const unattributed = slideInfo.filter((s) => !s.row && !s.template);
record(
  "every slide is either a merged row or a template slide",
  unattributed.length === 0 && mergedInfo.length > 0,
  unattributed.length
    ? `${unattributed.length} slide(s) name no row and hold no placeholder: ${unattributed.map(short).join(", ")}`
    : mergedInfo.length === 0
      ? "no merged slide found at all, so nothing below can be checked"
      : `merged ${mergedInfo.map((s) => `${short(s)}=${s.row.name}`).join(", ")}`,
);

/**
 * A copy that KEPT a placeholder, which the rule above cannot see.
 *
 * `TEMPLATE_MARK` cannot tell "a template slide" from "a copy the merge did not
 * finish" — both hold a placeholder for a field that has a column. So a slide
 * the merge half-filled is filed as TEMPLATE and skipped by every per-row check
 * below, which is the one defect this deck exists to catch being made invisible
 * by the fix that stopped template slides being marked down.
 *
 * Measured rather than argued: putting `{{Region}}` back on one merged copy of a
 * correct 13/13 round deck takes it to 8/13, and not one of the five failures
 * says a placeholder survived. They are the part counts and the formats, because
 * that slide left the merged population and took its chart, workbook and photo
 * with it. The deck is reported as broken, with the wrong reason, and a reader
 * goes looking for a missing chart part.
 *
 * The population that can answer is the slides that NAME A ROW: a template slide
 * holds `{{Name}}`, not `Ada`, so it names none and is not caught by this. It is
 * deliberately taken from `slideInfo` rather than `mergedInfo` — reading it off
 * the filtered list is what would make the check vacuous, since the filter is
 * what removed the slide.
 */
const halfFilled = slideInfo.filter((s) => s.row && s.template);
record(
  "no merged copy kept a placeholder the merge should have filled",
  halfFilled.length === 0,
  halfFilled.length
    ? `${halfFilled.length} slide(s) name a row AND still hold a placeholder: ${halfFilled
        .map((s) => `${short(s)}=${s.row.name} kept ${(s.text.match(TEMPLATE_MARKS) ?? []).join(" ")}`)
        .join(", ")}`
    : `${mergedInfo.length} merged slide(s), none holding a placeholder other than {{Nickname}}`,
);

// ------------------------------------------- which deck IS this

/**
 * The kit has two decks and this checker only ever knew one of them.
 *
 * Pointed at `modern-chart.pptx` it reported 4/13 — not because anything was
 * wrong with the deck, but because every check below is about the MAIN
 * template's shape: three classic charts, three photos, notes pages, a
 * `{{Nickname}}` that has to survive. The sunburst deck has none of those. The
 * round of 2026-08-30 read that 4/13 correctly as "wrong tool" and verified the
 * sunburst by hand, which is not something the next round should have to work
 * out for itself. A red that means nothing is worse than no red.
 *
 * A modern chart is not a `<c:chartSpace>` at all — PowerPoint stores it as a
 * chartEx part under its own relationship — so the two shapes are told apart by
 * which kind of chart part the deck actually holds.
 */
const chartExParts = names.filter((n) => /^ppt\/charts\/chartEx\d+\.xml$/.test(n)).sort();
const classicChartParts = names.filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n)).sort();
const deckKind = classicChartParts.length > 0 ? "kit" : chartExParts.length > 0 ? "modern" : "unknown";
console.log(`deck kind   : ${deckKind}\n`);

if (deckKind === "unknown") {
  // Said as a refusal rather than as a dozen failures. This checker knows two
  // decks; anything else gets an answer it can act on instead of a tally that
  // reads like a product defect.
  record(
    "this is a deck the checker knows how to check",
    false,
    "no classic chart and no chartEx part: this is neither the kit's main template " +
      "nor modern-chart.pptx, and every per-row check below is about one of those two. " +
      "Nothing is being claimed about this file.",
  );
  report();
}

if (deckKind === "modern") {
  // ----------------------------------------- the sunburst deck's own checks
  //
  // What the manual asks a human to look at, read off the bytes: each copy's
  // chart titled for its own row, the outer ring's FIRST segment merged and the
  // rest left alone, the inner ring untouched, and a fallback picture per copy
  // rather than the template's under somebody else's name.
  const charts = [];
  for (const s of mergedInfo) {
    for (const r of s.rels.filter((x) => /chart/i.test(x.type))) {
      const p = resolveTarget(s.path, r.target);
      if (/chartEx\d+\.xml$/.test(p)) charts.push({ slide: s, path: p, xml: await zip.file(p)?.async("string") });
    }
  }

  record(
    "one chartEx part per merged row",
    charts.length === mergedInfo.length && charts.length > 0,
    `${charts.length} chart(s) for ${mergedInfo.length} merged slide(s): ${charts.map((c) => c.path.replace("ppt/charts/", "")).join(", ") || "none"}`,
  );

  {
    const detail = [];
    let ok = charts.length > 0;
    for (const c of charts) {
      const text = textOf(c.xml ?? "");
      const own = text.includes(c.slide.row.name);
      const placeholder = /\{\{Name\}\}/.test(text);
      const foreign = ROWS.filter((r) => r !== c.slide.row && text.includes(`${r.name} pipeline`)).map((r) => r.name);
      if (!own || placeholder || foreign.length) ok = false;
      detail.push(
        `${c.path.replace("ppt/charts/", "")} own=${own ? c.slide.row.name : "NO"}` +
          (placeholder ? " PLACEHOLDER-LEFT" : "") +
          (foreign.length ? ` FOREIGN=${foreign.join("/")}` : ""),
      );
    }
    record("each chart's title names its OWN row", ok, detail.join(" | ") || "no chartEx found on any merged slide");
  }

  {
    // The outer ring's first cell is the placeholder; the two beside it are
    // plain text in the template and must come through untouched. All three
    // changing is a merge writing where no placeholder was — the failure this
    // deck exists to catch.
    const detail = [];
    let ok = charts.length > 0;
    for (const c of charts) {
      const pts = [...(c.xml ?? "").matchAll(/<cx:pt[^>]*>(.*?)<\/cx:pt>/gs)].map((m) => m[1].trim());
      const first = pts[0] ?? "";
      const merged = first === c.slide.row.region;
      const untouched = pts.includes("Benelux") && pts.includes("DACH");
      const placeholder = pts.includes("{{Region}}");
      if (!merged || !untouched || placeholder) ok = false;
      detail.push(
        `${c.path.replace("ppt/charts/", "")} first=${first || "(none)"}${merged ? "" : ` WANTED ${c.slide.row.region}`}` +
          (untouched ? "" : " OUTER-RING-OVERWRITTEN") +
          (placeholder ? " PLACEHOLDER-LEFT" : ""),
      );
    }
    record("the outer ring's FIRST segment is merged and the others are not", ok, detail.join(" | ") || "no chartEx");
  }

  {
    const detail = [];
    let ok = charts.length > 0;
    for (const c of charts) {
      const pts = [...(c.xml ?? "").matchAll(/<cx:pt[^>]*>(.*?)<\/cx:pt>/gs)].map((m) => m[1].trim());
      const kept = pts.includes("Existing") && pts.includes("New");
      if (!kept) ok = false;
      detail.push(`${c.path.replace("ppt/charts/", "")} ${kept ? "Existing+New kept" : "INNER RING LOST"}`);
    }
    record("the inner ring still reads Existing and New", ok, detail.join(" | ") || "no chartEx");
  }

  {
    // A modern chart carries a rendered picture for hosts that cannot draw it.
    // If the copies share one, every recipient is looking at the TEMPLATE's
    // figures under their own name — which no other check here would notice,
    // because the chart data itself is right.
    const byslide = [];
    for (const s of mergedInfo) {
      for (const r of s.rels.filter((x) => x.type.endsWith("/image") && !x.external)) {
        const p = resolveTarget(s.path, r.target);
        const buf = await zip.file(p)?.async("nodebuffer");
        if (buf) byslide.push({ slide: short(s), part: p.replace("ppt/media/", ""), sha: sha(buf) });
      }
    }
    const distinct = new Set(byslide.map((b) => b.sha));
    record(
      "each copy carries its OWN fallback picture, not the template's",
      byslide.length > 0 && distinct.size === byslide.length,
      byslide.length === 0
        ? "no fallback picture on any merged slide"
        : `${distinct.size} distinct across ${byslide.length}: ${byslide.map((b) => `${b.slide}->${b.part} ${b.sha}`).join(" | ")}`,
    );
  }

  report();
}

// ------------------------------------------- 2. part counts, over the copies

{
  const charts = new Set();
  const workbooks = new Set();
  const media = new Set();
  for (const s of mergedInfo) {
    for (const r of s.rels) {
      const p = resolveTarget(s.path, r.target);
      if (r.type.endsWith("/chart")) {
        charts.add(p);
        for (const wr of await readRels(zip, p)) {
          if (wr.type.endsWith("/package") || /\.xlsx$/.test(wr.target)) workbooks.add(resolveTarget(p, wr.target));
        }
      }
      if (r.type.endsWith("/image") && !r.external) media.add(p);
    }
  }
  const list = (s) => [...s].sort().join(", ") || "none";
  record("three chart parts, one per merged row", charts.size === 3, `${charts.size}: ${list(charts)}`);
  record("three embedded workbooks, one per merged row", workbooks.size === 3, `${workbooks.size}: ${list(workbooks)}`);
  record("three media parts, one per merged row", media.size === 3, `${media.size}: ${list(media)}`);
}

// ------------------------------------------- 3. charts: title + strCache

{
  const detail = [];
  let ok = true;
  let seen = 0;
  for (const s of mergedInfo) {
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
  for (const s of mergedInfo) {
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
  for (const s of mergedInfo) {
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
  // Counted over the COPIES. The template slide it came from carries it too, so
  // asking the whole deck for exactly three found five on a correct round and
  // called it a failure.
  const withNickname = mergedInfo.filter((s) => s.text.includes("{{Nickname}}"));
  record(
    "{{Nickname}} still present on the copies (no such column, must not blank out)",
    withNickname.length === 3,
    `${withNickname.length} of ${mergedInfo.length} merged slide(s) carry it: ${withNickname.map(short).join(", ") || "none"}`,
  );
}

// ------------------------------------------- 7. notes pages

{
  const detail = [];
  let ok = true;
  let seen = 0;
  for (const s of mergedInfo) {
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
  for (const s of mergedInfo) {
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
  for (const s of mergedInfo) {
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
    for (const s of mergedInfo) {
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

report();
