#!/usr/bin/env node
/**
 * Watch the sibling project for findings nobody here has answered.
 *
 * Most of what this repo knows about the PowerPoint host was learned by
 * PowerChart over its real-host rounds, and until now it arrived by somebody
 * happening to read that repo. That is not a process; it is luck with a good
 * afternoon attached, and `docs/SIBLING.md` records what the luck already cost
 * — 44 hand-copied citations, five of them carrying a counter that goes stale
 * the next time a round runs.
 *
 * So: look every week, automatically, and say only what is NEW.
 *
 * **`TRIAGED` is the load-bearing half**, exactly as `KNOWN_ISSUES` is for
 * `office-js-watch.mjs`. Without it the sweep reports the same seventy findings
 * every Monday and is filtered within a month. With it the report is precisely
 * "here is something nobody here has looked at", and the table doubles as the
 * record of what was decided — which is the question a reader actually has when
 * they meet a finding referenced in a comment.
 *
 * Usage:
 *   node scripts/sibling-watch.mjs                  # fetch and report
 *   node scripts/sibling-watch.mjs --from dir/      # report from saved copies
 *   node scripts/sibling-watch.mjs --json           # machine-readable output
 *
 * Exit 0 when everything is answered, 3 when something is not, anything else
 * when the sweep itself broke. A broken run must never read as a quiet week.
 *
 * **Raw file reads only, never the GitHub API.** Both repositories are public,
 * so `raw.githubusercontent.com` answers without a token — which means this
 * runs in CI and in an agent session alike. `office-js-watch.mjs` reaches the
 * office-js API and says in its own docstring that the fetch half cannot run
 * everywhere; this one has no such half, deliberately.
 */
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { isMain } from "./is-main.mjs";

const RAW = "https://raw.githubusercontent.com/dannbleeker/powerchart/main";

/**
 * The sibling's curated tables, and nothing else.
 *
 * Its `CLAUDE.md` is the fullest record and the worst feed: over a thousand
 * lines of prose that diffs into noise. These four files are keyed, small, and
 * already worded for a machine to compare — they are what that project uses to
 * gate its own fake against a real host.
 */
export const SOURCES = [
  { path: "scripts/office-js-watch.mjs", table: "KNOWN_ISSUES", kind: "issue" },
  { path: "scripts/host-baseline.mjs", table: "FAKE_BASELINE", kind: "question" },
  { path: "scripts/host-baseline.mjs", table: "KNOWN_DIVERGENCES", kind: "question" },
  { path: "scripts/host-baseline.mjs", table: "UNSTABLE_ANSWERS", kind: "question" },
  { path: "scripts/host-baseline.mjs", table: "PENDING_QUESTIONS", kind: "question" },
];

/**
 * The keys of one `export const NAME = { … }` table.
 *
 * A regex over two-space-indented keys rather than an import, because importing
 * would run the sibling's code — a weekly job that executes a file fetched over
 * the network is a supply chain, not a sweep. Both quoted and bare keys, since
 * that repo uses both.
 *
 * Returns null when the table is not there at all, which is a BROKEN SWEEP and
 * not an empty one. A rename upstream would otherwise report "nothing new"
 * forever, which is the failure this whole file exists to prevent, committed by
 * the file itself.
 *
 * **Null for a table that parses to NOTHING, too**, and that is the same rule
 * rather than a new one. The key pattern is anchored on two-space indentation,
 * so a reindent upstream — a formatter run, a nesting change — yields an empty
 * array, which is not null and therefore read as "nothing new". The protection
 * above covered a rename and not a reformat, which is the identical silence by
 * a likelier route. A curated table that genuinely holds nothing is not a state
 * worth distinguishing here: reporting it once as broken costs a person a
 * minute, and reading it as quiet costs every Monday after.
 *
 * The key charset takes `_` and `.` as well. It matched `[A-Za-z0-9-]` only, so
 * `brand_new_question` and `another.new.question` were dropped in silence —
 * from a sweep whose entire job is noticing something new. The sibling's ids
 * happen to use hyphens today, which is what made it latent.
 *
 * @param {string} source
 * @param {string} table
 * @returns {string[] | null}
 */
export function tableKeys(source, table) {
  const at = source.indexOf(`export const ${table}`);
  if (at < 0) return null;
  const rest = source.slice(at);
  const end = rest.search(/^\};/m);
  const body = end < 0 ? rest : rest.slice(0, end);
  const keys = [...body.matchAll(/^ {2}"?([A-Za-z0-9][A-Za-z0-9._-]*)"?:/gm)].map((m) => String(m[1]));
  return keys.length === 0 ? null : keys;
}

/**
 * What this repo has answered about the sibling's findings.
 *
 * Every key carries its reason, and **"no exposure" is a real answer** — the
 * most common one worth writing down, because an untriaged finding is
 * indistinguishable from an unnoticed one. Seeded by running this sweep against
 * the live tables and triaging what it reported, so the first Monday is quiet
 * rather than a wall of seventy rows nobody reads.
 *
 * Keys are `issue:<number>` and `question:<id>`, matching the sibling's own
 * spelling so a rename there surfaces as a new finding rather than silently
 * matching nothing.
 */
export const TRIAGED = {
  "issue:1650":
    "ADOPTED — a slide add whose sync never resolves though the slide lands. ADOPTED as doctrine rather than as a guard for a call we make: we never call `slides.add`, but `insertSlidesFromBase64` is the same shape, and the answer is the same one — the deck DELTA is the evidence, never the absence of an error. `insertDeck` counts before and after.",
  "issue:2328":
    "NO EXPOSURE — slideMaster.shapes throws GeneralException on the web. NO EXPOSURE — nothing here reads a master. Our cloned slides carry their layout relationship inside the package; the API is never asked.",
  "issue:2699":
    "NO EXPOSURE — a blank slide that is only blank to the eye. NO EXPOSURE — nothing here inspects whether a slide is blank.",
  "issue:2775":
    "RELEVANT — addTextBox deletes the SELECTED shape, web only. NO EXPOSURE to the call — this add-in never adds a text box. The CLASS matters though: the preview inserts at a moment when the user may have something selected, so `insertWhileSelectedProbe` asks whether an insert survives a standing selection rather than assuming it does.",
  "issue:2172":
    "NO EXPOSURE — addGeometricShape refused on a completely blank slide. NO EXPOSURE — no shape is ever added through the API.",
  "issue:2780":
    "NO EXPOSURE — a documented caveat with no code depending on it, in the sibling and here alike. NO EXPOSURE.",
  "issue:2881":
    "NO EXPOSURE — complex SVG renders wrong through the picture path. NO EXPOSURE — nothing here inserts a picture.",
  "issue:2903":
    "ADOPTED — a stale shape proxy answers `InvalidParam passed to GetItem(id)`, and that is the architecture here: it is the reason merge metadata is written into `ppt/tags/tagN.xml` before the insert rather than through a proxy afterwards. Its wait-two-seconds workaround is separately REJECTED — see the backlog.",
  "issue:3014": "NO EXPOSURE — powerPoint's API has no grouping story. NO EXPOSURE — no shape-level work at all.",
  "issue:3083":
    "NO EXPOSURE — setSelectedShapes([]) does not clear the selection on the web. NO EXPOSURE — this add-in never calls `setSelectedShapes`, which is what puts `getSelectedSlides` on the safe part of that surface.",
  "issue:3269":
    "NO EXPOSURE — office.js cannot read speaker notes at all. NO EXPOSURE, and worth saying why rather than leaving it to the label: `clone.ts` gives every copied slide its own notes page, which happens INSIDE the package. The API gap is real and we never reach for the API.",
  "issue:3309": "NO EXPOSURE — sVG cannot be read back out of a shape. NO EXPOSURE.",
  "issue:3698":
    "RELEVANT — a picture cannot be inserted while a shape is selected, and setSelectedShapes([]) may never resolve. NO EXPOSURE to the picture half. The selection half is why `insertWhileSelectedProbe` exists: we insert SLIDES, which is neither case, and nobody has established that it is safe.",
  "issue:3826":
    "NO EXPOSURE — a freshly-added slide's layout shapes throw GeneralException. NO EXPOSURE — nothing here reads a slide's layout through the API.",
  "issue:4272":
    "ADOPTED — a collection load of more than ~50 items answers short. ADOPTED — `ID_PAGE = 20` in `src/office/powerpoint.ts`, the same number for the same reason, and asked directly by the `deckRead` probe question.",
  "issue:4906":
    "NO EXPOSURE — slideLayout.shapes throws on decks built from a custom template. NO EXPOSURE — see #3826.",
  "issue:5455":
    "NO EXPOSURE — generalException reading ParagraphFormat.horizontalAlignment. NO EXPOSURE — text is replaced in the package, never through a text range.",
  "issue:6079":
    "RELEVANT — powerPoint on the WEB uppercases tag keys internally and then requires the uppercased spelling to read them back. RELEVANT AND ALREADY SAFE, by luck rather than by design until now: every key this engine writes (`SSF_MERGE_RUN`, `SSF_MERGE_RECORD`, `SSF_MERGE_BLOCK`) is already uppercase. `test/pptx.test.ts` should keep it that way — a lowercase key would be written into the package fine and be unreadable on the web.",
  "issue:2474":
    "ADOPTED — `SlideRange.id` lacks the `#XYZ` suffix the deck's own list carries. ADOPTED — `deckIdForSelectedSlide` matches by prefix and refuses two matches rather than guessing.",
  "issue:3565":
    "NO EXPOSURE — context.sync taking progressively longer every run. NO EXPOSURE, and the sibling's own note says why: it is WORD FOR MAC, not PowerPoint web, so it is not evidence about this host either.",
  "issue:6867":
    "RELEVANT — `Slide.exportAsBase64` omits modern comments and `ppt/authors.xml` from the exported deck. **A FINDING FOR US THAT THE SIBLING CORRECTLY MARKED NO EXPOSURE.** It calls that API for a PICTURE of a slide; we call `exportAsBase64Presentation` to read the TEMPLATE we then clone, so a part the export drops is a part the merged slides never get. Different call in the same family, and nobody has checked whether the presentation-level export drops the same parts. Open: worth a probe question before the first real merge on a deck with comments.",
  "issue:3784":
    "RELEVANT — shape TAGS are lost when a shape is cut and pasted on PowerPoint web. RELEVANT and documented rather than guarded: a merged slide cut and pasted into another deck loses its run tag, so undo cannot find it. Already carried as a caveat in `docs/MANUAL.md`. Detecting it is refused for the sibling's reason — the count would be swamped by tag writes the host refuses anyway.",
  "issue:6266":
    "NO EXPOSURE — getImageAsBase64 differs between Mac and Windows for content add-ins. NO EXPOSURE — nothing here rasterises anything.",
  "issue:6498":
    "RELEVANT — shapes inserted on the web may appear in the slide PREVIEW but not the main view without a refresh. RELEVANT as a support answer rather than as a defect to fix: a user reporting that merged slides are missing may be seeing this, and the deck delta will say the slides landed. Nothing here can read the canvas to tell them apart.",
  "issue:5022":
    "NO EXPOSURE — context.sync runs indefinitely when shapes are re-read after an image insert. NO EXPOSURE — no image inserts and no shape re-reads.",
  "issue:5101":
    "NO EXPOSURE — a placeholder keeps `type: Placeholder` when reused. NO EXPOSURE — nothing here reads a shape's type.",
  "issue:5264":
    "NO EXPOSURE — a part of the object model Office.js cannot reach. NO EXPOSURE — recorded as a limitation on both sides.",
  "issue:5849": "NO EXPOSURE — shape.group throws GeneralException. NO EXPOSURE — no grouping.",
  "issue:5896": "NO EXPOSURE — reported alongside another SVG defect. NO EXPOSURE.",
  "issue:6363":
    "RELEVANT — `PowerPoint.run`'s batching fails to load properties reliably — properties not available after `context.sync()`, web only. HIGHLY RELEVANT: `deckSlideIds` batches a `load(\"id\")` across up to twenty `getItemAt` handles and reads them all after one sync, which is precisely the shape this describes. Asked by the `deckRead` probe question's `empty` arm.",
  "issue:2714": "NO EXPOSURE — setSelectedDataAsync converts points to pixels. NO EXPOSURE — never called.",
  "question:getcount-populates-same-sync":
    "RELEVANT: `deckSlideIds` calls `getCount()` and then reads `getItemAt` handles, and the sibling has recorded a host whose count is right while the list is empty. This is why the paging loop trusts the scalar count and not a collection load.",
  "question:getitemat-past-end":
    "RELEVANT: both `deckSlideIds` (paging by index) and `undoInsert` (deleting by index, highest first) index into the collection. What the host does past the end bounds both.",
  "question:which-end-a-short-read-drops":
    "RELEVANT: office-js#4272 again. Our own probe asks the same thing as `prefixOk` — whether a short read is the first n IN DECK ORDER — because a short read that is not a prefix makes a slide NUMBER wrong rather than merely a list shorter.",
  "question:how-many-collection-reads-a-context-survives":
    "RELEVANT: `deckSlideIds` makes one `PowerPoint.run` per page, deliberately, so no context accumulates reads. Recorded so the reason survives if anyone tries to make it one batch.",
  "question:delete-then-lookup":
    "RELEVANT as doctrine: undo deletes by position, highest index first, and re-counts the deck rather than believing the call. Whether a deleted slide still resolves is exactly the question that made by-id clean-up unsafe.",
  "question:scratch-slides-returned":
    "ADOPTED as doctrine. The sibling's by-id clean-up left 45 blank slides; our probe's sweep is positional and triple-clamped, and each clamp is proven load-bearing in `test/undo.test.ts`.",
  "question:shapes-by-index-vs-items": "NO EXPOSURE — a question about a SHAPE collection. Nothing here reads shapes.",
  "question:creationid-on-fresh-shape":
    "NO EXPOSURE, and worth saying why: we do write `p14:creationId`, but on a cloned SLIDE and inside the package. These three questions are about a shape's creation id read back through the API.",
  "question:creationid-survives-a-sync": "NO EXPOSURE — see `creationid-on-fresh-shape`.",
  "question:creationid-survives-grouping":
    "NO EXPOSURE — see `creationid-on-fresh-shape`; there is no grouping here either.",
  "question:load-isnullobject-populates":
    "NO EXPOSURE — this add-in never calls `getItemOrNullObject`. Slides are reached by `getItemAt`, which is a different code path.",
  "question:load-id-populates-isnullobject": "NO EXPOSURE — see `load-isnullobject-populates`.",
  "question:getitemornullobject-missing": "NO EXPOSURE — see `load-isnullobject-populates`.",
  "question:shape-add-fresh-slide-proxy": "NO EXPOSURE — adds a shape to a slide proxy. No shape work here.",
  "question:shape-add-held-slide-proxy":
    "NO EXPOSURE — see `shape-add-fresh-slide-proxy`. Also in UNSTABLE_ANSWERS, so nothing may be built on its answer in either repo.",
  "question:shape-add-held-slide-proxy-again": "NO EXPOSURE — the partner of `shape-add-held-slide-proxy`.",
  "question:shape-resolve-held-slide-proxy": "NO EXPOSURE — resolving a held shape proxy.",
  "question:shape-add-fresh-getitem-slide": "NO EXPOSURE — see `shape-add-fresh-slide-proxy`.",
  "question:shape-add-positional-slide-proxy": "NO EXPOSURE — see `shape-add-fresh-slide-proxy`.",
  "question:shape-proxy-survives-one-sync":
    "NO EXPOSURE — no proxy is held across a sync here; a merge holds only slide indices.",
  "question:shapes-items-count-honest":
    "NO EXPOSURE to the SHAPE collection. The equivalent question about the SLIDE collection is the one that matters here, and our own probe asks it as `deckRead`.",
  "question:shapes-items-via-positional-slide": "NO EXPOSURE — see `shapes-items-count-honest`.",
  "question:tags-add-same-key-twice":
    "NO EXPOSURE to the API question. The package equivalent is real and already handled: `tags.ts` finds the next free `tagN` rather than overwriting `tag1.xml`, which would destroy another tool's tags.",
  "question:tags-on-fresh-shape":
    "NO EXPOSURE — tags are written into the package, never onto a shape through the API.",
  "question:tag-through-refetched-shape": "NO EXPOSURE — see `tags-on-fresh-shape`.",
  "question:how-many-syncs-a-creation-handle-survives":
    "NO EXPOSURE — no handle from a creation call is ever held here; an insert returns nothing we keep.",
  "question:collection-read-poisons-the-creation-handle":
    "NO EXPOSURE — see `how-many-syncs-a-creation-handle-survives`.",
  "question:does-a-failed-group-poison-the-tag": "NO EXPOSURE — no grouping.",
  "question:addgroup-returns-usable": "NO EXPOSURE — no grouping.",
  "question:group-children-via-getcount": "NO EXPOSURE — no grouping.",
  "question:group-reports-its-children": "NO EXPOSURE — no grouping.",
  "question:group-of-existing-shape-readable": "NO EXPOSURE — no grouping.",
  "question:binding-names-shape-later":
    "NO EXPOSURE, and the route is REJECTED rather than merely unused: the sibling's probe showed the host rejects the batch carrying the binding, with a control arm proving it was the binding. See the backlog's rejected list.",
  "question:picture-then-shape-read": "NO EXPOSURE — no picture inserts.",
  "question:slide-layout-readable": "NO EXPOSURE — nothing here reads a layout through the API.",
  "question:layouts-readable": "NO EXPOSURE — see `slide-layout-readable`.",
  "question:untrack-available":
    "NO EXPOSURE — a merge holds a handful of proxies for one batch, so there is nothing to untrack. The sibling measured it unavailable on this host anyway.",
  "question:untrack-available-on-shape": "NO EXPOSURE — see `untrack-available`.",
};

/**
 * The closed vocabulary every `TRIAGED` reason opens with.
 *
 * A prefix rather than prose, because something has to be able to ASK which
 * findings we acted on. `docs/SIBLING.md` is the human record and it can fall
 * behind this table silently — which is the exact drift this whole apparatus
 * exists to stop, and it would be embarrassing to build it with the flaw
 * inside. `test/sibling.test.ts` asserts the ledger mentions every finding that
 * is not NO EXPOSURE.
 *
 * `NO EXPOSURE` is first because it is the answer nobody writes down and the
 * one that makes the sweep quiet enough to read.
 */
export const VERDICTS = ["NO EXPOSURE", "ADOPTED", "RELEVANT"];

/**
 * The verdict a reason opens with, or undefined when it opens with none.
 *
 * @param {unknown} reason
 * @returns {string | undefined}
 */
export function verdictOf(reason) {
  return VERDICTS.find((v) => typeof reason === "string" && reason.startsWith(v));
}

/**
 * Findings present upstream with no row in `TRIAGED`.
 *
 * @param {{kind: string, table: string, keys: string[]}[]} tables
 * @param {Record<string, string>} [triaged]
 * @returns {{id: string, kind: string, key: string, tables: string[]}[]}
 */
export function untriaged(tables, triaged = TRIAGED) {
  /** @type {Map<string, {id: string, kind: string, key: string, tables: string[]}>} */
  const seen = new Map();
  for (const { kind, table, keys } of tables) {
    for (const key of keys) {
      const id = `${kind}:${key}`;
      if (Object.prototype.hasOwnProperty.call(triaged, id)) continue;
      const hit = seen.get(id);
      if (hit) hit.tables.push(table);
      else seen.set(id, { id, kind, key, tables: [table] });
    }
  }
  return [...seen.values()];
}

/**
 * Every table, from whatever `read` hands back for a path.
 *
 * @param {(path: string) => string} read
 * @returns {{path: string, table: string, kind: string, keys: string[]}[]}
 */
export function tablesFrom(read) {
  /** @type {Map<string, string>} */
  const cache = new Map();
  /** @type {{path: string, table: string, kind: string, keys: string[]}[]} */
  const out = [];
  for (const { path, table, kind } of SOURCES) {
    if (!cache.has(path)) cache.set(path, read(path));
    const keys = tableKeys(cache.get(path), table);
    if (keys === null)
      throw new Error(
        `${path}: no keys read from ${table} — it was renamed, moved, emptied or reformatted upstream. ` +
          `A sweep that cannot read the table must say so rather than report a quiet week.`,
      );
    out.push({ path, table, kind, keys });
  }
  return out;
}

/**
 * @param {{id: string, kind: string, key: string, tables: string[]}[]} found
 * @returns {string}
 */
function report(found) {
  if (!found.length) return "Nothing new: every finding in the sibling's curated tables has a row in `TRIAGED`.\n";
  const issues = found.filter((f) => f.kind === "issue");
  const questions = found.filter((f) => f.kind === "question");
  const lines = [
    "## Sibling findings nobody here has answered",
    "",
    "Each needs a row in `TRIAGED` in `scripts/sibling-watch.mjs` and a line in",
    "`docs/SIBLING.md`. **“No exposure” is a real answer** — say it rather than",
    "leaving the row out, or this comes back next Monday looking identical.",
    "",
  ];
  if (issues.length) {
    lines.push(`### office-js issues (${issues.length})`, "");
    for (const f of issues) lines.push(`- [#${f.key}](https://github.com/OfficeDev/office-js/issues/${f.key})`);
    lines.push("");
  }
  if (questions.length) {
    lines.push(`### Host questions (${questions.length})`, "");
    // Which table a question sits in IS information: `UNSTABLE_ANSWERS` means
    // the sibling has seen it answer two ways, so building on it is a mistake
    // whatever it says today.
    for (const f of questions) lines.push(`- \`${f.key}\` — ${f.tables.join(", ")}`);
    lines.push("");
  }
  return lines.join("\n");
}

async function main(argv) {
  const json = argv.includes("--json");
  const fromAt = argv.indexOf("--from");
  const dir = fromAt >= 0 ? argv[fromAt + 1] : undefined;
  const read = dir
    ? (/** @type {string} */ path) => readFileSync(join(dir, String(path.split("/").pop())), "utf8")
    : await (async () => {
        /** @type {Map<string, string>} */
        const bodies = new Map();
        for (const path of [...new Set(SOURCES.map((s) => s.path))]) {
          const res = await fetch(`${RAW}/${path}`);
          if (!res.ok) throw new Error(`${path}: the sibling repo answered ${res.status}`);
          bodies.set(path, await res.text());
        }
        return (/** @type {string} */ path) => String(bodies.get(path));
      })();

  const tables = tablesFrom(read);
  const found = untriaged(tables);
  process.stdout.write(
    json
      ? `${JSON.stringify({ found, tables: tables.map((t) => ({ ...t, keys: t.keys.length })) }, null, 2)}\n`
      : report(found),
  );
  return found.length ? 3 : 0;
}

if (isMain(import.meta.url)) {
  main(process.argv.slice(2)).then(
    (code) => process.exit(code),
    (err) => {
      process.stderr.write(`sibling-watch: ${err instanceof Error ? err.message : String(err)}\n`);
      process.exit(1);
    },
  );
}
