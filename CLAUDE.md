# SSF Merge — project memory

Mail merge for PowerPoint, as an Office.js task pane add-in. Part of the SSF
add-in family. Sibling project: PowerChart, which is where most of what this
repo knows about the PowerPoint host was learned — `docs/SIBLING.md` is the
ledger of what came across, what was done about each item, and the rules that
keep a borrowed fact from going quietly bad.

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
| `src/host/` | the DECISIONS about talking to a host, all pure and all tested: `capability.ts` (version floor, where the template's bytes come from), `verdicts.ts` (what an observation means), `undo.ts` (which slides a run may take back), `timeout.ts` (what each call is allowed to cost) |
| `src/office/` | the Office.js CALLS, and nothing else. Every judgement is imported from `src/host` |
| `src/pane/` | `steps.ts` (which step is reachable, what the one button says), `summary.ts` (the sentences numbers go into), `render.ts` (the DOM), `main.ts` (**the only file here allowed to touch Office.js**), plus the HTML and the SSF stylesheet |
| `test/fixtures/` | `deck.ts` builds a minimal .pptx in memory, so no test depends on a committed binary |

**`src/host` decides, `src/office` calls, and `test/architecture.test.ts` holds
both directions.** An Office.js import in `src/host` makes a rule untestable; a
rule reimplemented inline in `src/office` looks tidier and rots quietly, so every
file there must import from `src/host`. This is the probe's split — dumb snippet,
tested verdicts — applied to the product, and it is why three sheets could be
read at all.

`src/office` is **not** in the coverage include list, deliberately. Pooling a
well-tested engine with untestable host calls produces one number that hides
both.

**The pane is the one surface the suite cannot judge.** `pane-render.test.ts`
pins its behaviour in jsdom, which has no layout and no colour, so a rule about
how it LOOKS is invisible there unless somebody turns it into something
countable. `scripts/pane-shots.mjs` renders every state at 320 and 512 — the ends
of the width a task pane is dragged between — and looking at the output is part
of done:

```bash
npx vite --port 5199 --strictPort &
node scripts/pane-shots.mjs          # PNGs in /tmp/pane-shots
```

Its first run found a real defect no test had: the fields step drew the orange
tick AND an orange-bordered chip, breaking the **orange budget** the layout was
approved on — one orange element per view. The fix was to make the budget a
single function (`orangeHolder`) rather than a condition at each call site, and
the finding is now a test that counts orange-carrying classes across every state
and step. **A rule enforced in one place can be tested; one enforced in three
cannot.**

Proving that guard took two attempts, and the first is the lesson: removing the
`unmatched` branch turned a different assertion red, because it stopped the chip
being orange rather than putting a second orange on screen. Only a faithful
revert of the RENDER change reproduced the defect. Check which assertion goes
red, not just that one does.

## What THIS host answered

Measured on PowerPoint for the web, 2026-08-26, two sheets under
`docs/host-answers/`. Recordings, not opinions.

- **The package path works.** A cloned slide with a fresh creation id inserts.
  So does the presentation's own bytes, so does the same package under
  `UseDestinationTheme`. This was the assumption the whole architecture rested
  on and it is now measured rather than hoped for.
- **A tag written into `ppt/tags/tagN.xml` reads back through
  `slide.tags.getItemOrNullObject`.** It came back holding `probe-run`. Merge
  metadata can go into the file before the insert and be read from the host
  afterwards, which is the single largest risk this project had.
- **office-js#6105 does not reproduce here.** The collision arm — two slides
  deliberately sharing one creation id — inserted cleanly. Keep the rewrite
  anyway: the hosts that need it are the ones nobody here can test.
- **A shape proxy does not survive a `context.sync()`.** `5010: InvalidParam
  passed to GetItem(id)`, raised the first time a shape created a sync earlier
  was touched again. Office.js rewrites a created shape's object path to
  `shapes.getItem(id)` once it has been through a sync, and this host refuses
  that id — the shape is there, it just will not be named again. **Queue every
  write against a shape in the batch that created it.** There is no way round
  it by name: `ShapeCollection.getItem` is documented in the typings as taking
  an ID, not a name, whatever the docs article's example suggests.
- **A targeted substring write keeps the formatting around it.** `getSubstring(a,
  n).text = v` replaced a placeholder and the bold on that run survived. Live
  preview can target a substring; it does not have to redraw whole shapes.
- **Two writes queued in one batch DO interfere: the second sees the first
  one's result.** Writing five characters over `AAA` in `AAA-BBB` and then three
  at the original offset of `BBB` gave `XXXX2BB`, not `XXXXX-2`. **So Office.js
  replacements must be queued RIGHT TO LEFT**, highest offset first, or every
  placeholder after the first lands in the wrong place — and it lands in a way
  that reads as a data bug rather than an ordering bug.

  This constrains the host/preview layer only. `mergeParagraph` writes character
  buffers built from the original joined text and applies every hit against
  those, so the package path is immune by construction. Do not "fix" it to match
  this finding.
- **A call can raise and still have done the work.** The third sheet's 30-second
  budget expired on an insert whose deck delta was exactly the two slides asked
  for. Read the DELTA, never the presence of an error: reading the error as
  decisive turned one late answer into three false statements in the same run.
  This is the opposite direction from PowerChart's "a stall is death, not
  slowness", and it is this host's own evidence rather than that one's.
- **The first sheet said none of this**, and the reason is worth keeping. Every
  insert answered `InvalidArgument` because our own fixture was malformed — an
  empty `<a:themeElements/>` where `CT_BaseStyles` requires three children. Two
  arms were added rather than reasoned about, and the second sheet answered
  everything. The blame arm is what turned "InvalidArgument" from a fact about
  the host into a fact about us.

  Three things changed between the sheets — the theme, the `docProps` parts and
  the text box's `xfrm` — so **the theme is not proven to be the whole cause**,
  only the one that violated a schema. Do not write it down as settled.

## Host rules, learned the expensive way

These come from PowerChart's field rounds against PowerPoint on the web. They
are recordings, not opinions. Do not re-derive them — and do not re-copy them
either: `docs/SIBLING.md` carries each one with its source and what was done
about it, including the findings that turned out to be **no exposure** here,
which is a real answer and the one nobody writes down.

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
  reply of every session in this repo, and for every reply after it. It is not a
  mode that gets switched on when somebody asks: it is on by default here, it
  does not lapse because a turn is long or a finding is interesting, and it does
  not need re-requesting. Only "stop caveman" turns it off, for that session.

  Code, commits, PR bodies and user-facing docs stay normal prose — they have
  readers who were not in the conversation.

  Drop it for security warnings, destructive-action confirmations, and any
  multi-step sequence where dropping conjunctions could be misread. Resume
  straight after.
- **A regression test must be proven to fail without its fix.** Break the source,
  re-run, confirm the new test goes red for the RIGHT reason, restore. Both
  guards in `test/text.test.ts` and `test/pptx.test.ts` were proven this way.
- **A guard that goes red for the wrong reason is worse than no guard.** The
  no-Office-imports test first matched the word "Office.js" in the comments
  explaining why the engine avoids it, and failed on four correct files.
- **A number copied from a live counter carries the date it was taken.**
  Otherwise it is a claim that rots. Four places said "174 consecutive archived
  rounds" and one said "passed 174 of 174"; PowerChart's archive held exactly
  174 that morning, so all five were right, and round 175 makes every one of
  them wrong with nothing anywhere to say so. Dated, the same number is a
  recording — it never becomes false, only older, and a reader can judge that.
  Two corollaries: where the number adds nothing, prefer the phrasing that
  cannot decay ("every rung answered in every archived round" is both stronger
  and permanent); and never pair figures from two measurements, which is how
  `docs/BACKLOG.md` came to carry a count from the run that measured 29.2s
  beside a maximum of 31.1s from a later one. `docs/SIBLING.md` has the sort.
- **Test files are named by topic, never by increment.**
- **Compare a .pptx by its PARTS, never by the archive's bytes.** JSZip stamps an
  entry time whenever a file is written, so two builds of identical content hash
  differently — and the order writes happen in changes those stamps. Releasing a
  parsed part mid-run produced a different zip hash and identical contents
  across all 66 parts, which for ten minutes looked like the change altering
  output. Load both, walk the part names, hash name plus content.
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
