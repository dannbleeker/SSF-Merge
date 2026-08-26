# SSF Merge — project memory

Mail merge for PowerPoint, as an Office.js task pane add-in. Part of the SSF
add-in family. Sibling project: PowerChart, which is where most of what this
repo knows about the PowerPoint host was learned.

## Architecture in one paragraph

A .pptx is a zip of XML parts. The merge happens **in the file**: clone the
template block's slide parts once per record, replace placeholders at the run
level where the original `<a:rPr>` is sitting untouched, write the merge tags
into `ppt/tags/tagN.xml`, then hand the finished deck to PowerPoint in a single
`insertSlidesFromBase64` call. Nothing is drawn shape by shape, so none of the
per-shape failure surfaces exist.

`src/core/` is pure TypeScript with **zero Office imports**, enforced by
`test/architecture.test.ts`. It takes bytes and records and returns bytes, which
is what lets it run in the pane, in a CLI and in the suite with no PowerPoint.

## Where things live

| directory | what it owns |
| --- | --- |
| `src/core/pptx/` | `pkg.ts` (the zip, parts, rels, content types, slide ids), `clone.ts` (slide cloning), `tags.ts` (tags written into the file), `xml.ts` (one parser everywhere) |
| `src/core/merge/` | `text.ts` (run-aware replacement) |
| `src/core/data/` | `recordset.ts` (parsing, type detection), `format.ts` (numbers, dates, case) |
| `test/fixtures/` | `deck.ts` builds a minimal .pptx in memory, so no test depends on a committed binary |

## Host rules, learned the expensive way

These come from PowerChart's field rounds against PowerPoint on the web. They
are recordings, not opinions. Do not re-derive them.

- **A slide the run just added does not resolve by id.** `slides.getItem(id)`
  refuses it, and `deleteSlideByPosition`'s `indexOf(id) < 0` reads "not listed"
  as "already gone" — one clean-up reported 45 deletes and removed nothing. So:
  merge metadata goes into the package before the insert, and undo is
  **positional** with clamps, never by id.
- **Tag writes through a shape proxy are refused.** 46 `InvalidParam passed to
  GetItem(id)` in one run. Writing `ppt/tags/tagN.xml` cannot be refused,
  because nothing is asked. Find the next free `tagN`: overwriting `tag1.xml`
  destroys another tool's tags and every slide pointing at them.
- **Without `targetSlideId`, `insertSlidesFromBase64` inserts at the FRONT.** A
  real run put 37 generated slides ahead of the user's title slide.
- **A queued delete that raises nothing has not necessarily happened.** Adds,
  inserts and tag writes have all been accepted and not performed. Confirm every
  destructive step with a second read.
- **An empty collection read is not an empty slide.** Both signals agreed at
  zero and both were wrong. Never claim data loss from a read; an id refusal
  anywhere makes every id in that operation suspect.
- **Two `slides.add()` calls 0.4s apart killed the tab.** Twice. Slides arrive
  through one insert, never a loop of adds.
- **Do not wait after adding a slide.** A deliberate delay cost 18 of 19 probe
  answers in one round. No sleeps in the merge path.
- **A stall is death, not slowness.** Of 327 answered batches the slowest took
  31s against a 45s budget; seventeen abandoned calls never came back. Size the
  insert timeout per slide and treat expiry as final.
- **Probe for the method, do not trust the requirement set.** Microsoft's own
  docs disagree about whether `insertSlidesFromBase64` is 1.2 or 1.5.
- **A fallback reading must never be cached as a measurement.** Record how a
  value was obtained and let only an authoritative source stick.
- **`load("items")` does not load the items' properties.** Name them:
  `load("items/id")`.
- **Reading the selection and holding it are different things.** On the web
  `addTextBox` deletes the selected shape (office-js#2775) and
  `setSelectedShapes` wedges the selection subsystem (#3083, #3698). Click-to-bind
  reads, drops, then writes. Navigation uses `setSelectedSlides`; nothing calls
  `setSelectedShapes`.
- **Shape tags do not survive cut/paste on the web** (office-js#3784). A merged
  slide copied to another deck loses its run tag, so undo will not find it. Say
  so in the docs; do not try to detect it.
- **A custom XML part written at the package root is invisible to Office.js.**
  Relate it from `ppt/presentation.xml`.
- **Blob downloads from the task pane are blocked in WebView2**
  (office-js#1511). Output goes into the deck, into a new presentation, or
  through `Office.context.ui.openBrowserWindow`.

## The lockstep rule (CI-enforced)

Any feature change updates, in the same PR:

1. **`docs/MANUAL.md`** — `test/docs.test.ts` fails on a format or tag key the
   manual does not mention, and on a manual that stops marking unbuilt sections
   as planned.
2. **`CHANGELOG.md`** — under `## [Unreleased]`, in the language a user would
   use, not the language of the diff.
3. **`docs/BACKLOG.md`** — an item that shipped is REMOVED, not ticked. Anything
   still listed is genuinely not done. A rejected idea moves to the rejected
   list with the reason, so nobody re-proposes it.
4. **`README.md`** feature table.

## Conventions

- **Answer in caveman style** (the `caveman` skill, `full` level) from the first
  reply of every session in this repo. Code, commits, PR bodies and user-facing
  docs stay normal prose. Drop it for security warnings, destructive-action
  confirmations, and any multi-step sequence where dropping conjunctions could
  be misread.
- **A regression test must be proven to fail without its fix.** Break the source,
  re-run, confirm the new test goes red for the RIGHT reason, restore. Both
  guards in `test/text.test.ts` and `test/pptx.test.ts` were proven this way.
- **A guard that goes red for the wrong reason is worse than no guard.** The
  no-Office-imports test first matched the word "Office.js" in the comments
  explaining why the engine avoids it, and failed on four correct files.
- **Test files are named by topic, never by increment.**
- **A check whose answer depends on a build artifact is not a check.**
  `dist-lib/` is compiled output and is absent on a fresh checkout, so the two
  scripts that import it (`build-probe.mjs`, `read-answers.mjs`) type-resolve to
  `error` in CI and to real types on a machine that has run `npm run build:lib`.
  Type-aware lint therefore passed locally and failed on the same source in CI.
  The fix is to scope the affected rule off for exactly those files, not to make
  CI build first: building would only hide the non-determinism behind whichever
  order the steps happen to run in. Before trusting any local run of a check that
  reads `dist-lib/`, delete it and run again.
- **All sample data is invented.** The repo is public.
- **Merging to `main` is authorized.** Once CI is green on the exact pushed
  commit — verify the run's `head_sha` matches the branch head, because a
  force-push leaves older runs attached to the PR and they are not evidence
  about the head — squash-merge without asking. Granted by the owner on
  2026-08-26. It does not extend to merging somebody else's PR, to changing
  what `main` requires, or to a red or conflicted head.
- Branch flow: develop on the session's designated `claude/*` branch, one PR per
  increment. `main` requires a pull request and a passing `test` check, and the
  branch is deleted on merge, so the next increment starts with
  `git checkout -B <branch> origin/main` rather than reusing what is there.
  Reusing a branch across a squash merge is what forced a force-push once: the
  old commits survive, `git log` calls them unmerged because squashing hides
  ancestry, and only a content diff against `main` proves otherwise.

## Open questions for the real host

Milestone zero, and nothing should be built on a guess about any of them:

1. Does inserting a cloned slide with a **fresh** creation id avoid the
   `InvalidArgument` of office-js#6105?
2. Does `getSubstring(start, len).text = v` keep the run's font and bullet?
3. Within one batch, do substring writes shift the offsets of later writes?
4. Does `fill.setImage` stretch, or preserve aspect ratio?
