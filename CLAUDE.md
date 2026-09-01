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
| `src/core/trace.ts` | the run record — a capped array cleared per run, one clock origin, one formatter for the screen and the file. In `src/core` because the coverage `include` is a fixed list of three globs and a new top-level directory would be measured by nothing |
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
of the width a task pane is dragged between — in BOTH themes, and looking at the
output is part of done:

```bash
npx vite --port 5199 --strictPort &
node scripts/pane-shots.mjs          # PNGs in /tmp/pane-shots, plus an audit
```

**It measures three things as well as shooting them**, because each is a number a
reader cannot take off a PNG and each has produced a real defect. It prints its
findings and exits 1.

- **Horizontal overflow.** Every long string on that screen comes from outside
  it — a column header, a row's first cell, a file name, an error PowerPoint
  wrote. An unbroken header took the pane to 545px inside a 320px frame and a
  spaceless host error to 3751, with the one filled button off the side.
  `readable` caps an error at 400 characters and a cap is not a break: 400
  characters with no space in them is one word.
- **Text contrast, in both themes.** Blue is a BACKGROUND here — the header and
  the primary button, both carrying white text — and it was also the ink on
  chips, field tags and every secondary button. On the dark palette that ink is
  3.0:1, and "Remove these slides", the whole way back from a merge, was 2.93:1.
  Disabled controls are exempt, which is WCAG 1.4.3 and not a convenience.
- **Where the keyboard is**, which a screenshot cannot show at all — a focus
  ring is only on screen while something is focused. The stylesheet's focus rule
  named `input`, `textarea` and `select` and stopped, so every BUTTON fell back
  to Chrome's own ring: `rgb(16, 16, 16)`, near-black, on the dark theme's
  near-black pane, at 1.03:1. The sweep presses Tab once first, because Chrome
  matches `:focus-visible` on a programmatic `focus()` only when the last
  interaction was a keyboard one — without it every button reports no outline
  and the measurement answers "clean" because it never looked.

**A weekly job runs it** (`.github/workflows/pane-audit.yml`), because
"somebody has to remember" is how the first two findings got in. Not a required
check and deliberately not part of `test`: it needs a dev server and a browser,
and a minutes-long job in front of every merge is one that gets switched off
after the first bad week. It keeps the PNGs as an artifact on failure only.

**Two fixtures exist only to exercise the overflow half** — an unbroken column
header and a spaceless notice — and they are the difference between a gate and a
gate's name. Nothing else in the set has a word long enough to need
`overflow-wrap`: the pivot headers all carry spaces, so without them that half
reports clean against a stylesheet that has lost the rule. Check what a new
measurement MATCHES before trusting that it passes.

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

Measured on PowerPoint for the web, six sheets under `docs/host-answers/`.
Recordings, not opinions.

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
- **`ShapeFill.setImage` STRETCHES. It does not preserve aspect ratio, crop or
  letterbox.** Measured 2026-08-28 by filling three rectangles from one SQUARE
  card — 1:1, 2:1 wide, 1:2 tall — and looking at the slide. The circle on the
  card is round in the square box, a wide ellipse in the wide one and a tall
  ellipse in the tall one, and all four corner labels survive in every box. An
  ellipse with its corners intact is a stretch; a crop would have lost corners
  and a letterbox would have shown bars.

  The only question in this project that no API could answer, and the only one
  whose evidence is a screenshot rather than a number. `scripts/build-aspect-probe.mjs`
  rebuilds the instrument if it is ever needed again.

  **What it decides is the ROUTE for image fields, not just a workaround.**
  Through `setImage` the engine would have to letterbox every image itself —
  pad it to the target shape's aspect ratio so that the host's stretch becomes
  a no-op — which means reading each shape's `<a:ext>` and compositing before
  sending. Through the PACKAGE, which is what this add-in does with everything
  else, the fill mode is ours to write: `<a:blipFill>` takes `<a:stretch>` or an
  `<a:srcRect>` crop, so the host's behaviour is not imposed at all. The
  measurement is therefore an argument FOR the package route rather than a task
  list for the other one.

- **`exportAsBase64Presentation` DROPS comments and `ppt/authors.xml`.** The
  sixth sheet (2026-08-28) put this on a deck carrying four comments: four
  comment parts and an authors part went in, none came out. So
  [office-js#6867](https://github.com/officedev/office-js/issues/6867) reaches
  the presentation-level call, not just the per-slide one.

  Harmless in itself — the template slides stay in the user's deck untouched, and
  a reviewer's note is not content a merge should reproduce. What it exposed is
  that **the two template routes disagreed**: the subset route (1.10) produced
  comment-free clones because the host dropped them, while the file route (every
  host below that) copied the slide's rels wholesale and gave EVERY clone a
  relationship to the template's comment part — three slides, one
  `modernComment_101_AEAB9DA1.xml`, measured. A "check this with Legal" on all
  240 merged slides, as one shared thread.

  `cloneSlide` drops comment relationships now, in both spellings — the classic
  `commentN.xml` and the modern `modernComment_<id>_<hash>.xml` the web writes
  under a Microsoft namespace — and `removeSlide` takes a comment part away with
  its slide, as it already did for notes. The routes agree, and they agree on the
  answer the host had already chosen.

- **A collection load of the deck's slides answers IN FULL, past the ceiling.**
  The fifth sheet (2026-08-28) put this on a 58-slide deck — above the ~50 that
  [office-js#4272](https://github.com/officedev/office-js/issues/4272) describes,
  which is what makes the answer worth anything; the sheet before it asked on
  eight slides and could say nothing. `load("items/id")` answered all 58, in deck
  order, with nothing short and nothing empty. So
  [office-js#6363](https://github.com/officedev/office-js/issues/6363) does not
  reproduce here either.

  **`deckSlideIds` keeps paging anyway, and that is now a decision rather than a
  habit.** The evidence says one load would do on this host in that minute; the
  paging costs a few syncs and removes a failure mode nobody would recognise
  from the symptom. What the answer buys is knowing that a short read is not
  what to suspect first when something goes wrong here.
- **A slide insert survives a STANDING SELECTION.** Three shapes selected, both
  slides landed. office-js#2775 (a text-box add deletes the selected shape) and
  #3698 (a picture will not insert while one is selected) are about SHAPES, and
  this add-in inserts SLIDES — that was the reasoning, and it is measured now.
  **So nothing here needs `setSelectedShapes`**, which is the one call in this
  family with a history of wedging the host. Do not add a "drop the selection
  first" step; there is no defect to fix and the fix would be the dangerous part.
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
which is a real answer and the one nobody writes down. `npm run sibling-watch`
reports anything in the sibling's curated tables that has no row yet; it runs
weekly on its own and files one issue, reopened rather than duplicated.

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
- **The two template routes return DIFFERENT things, and only one of them is
  the block.** `exportAsBase64Presentation` (1.10 and up) hands back a package
  holding only the slides asked for; `getFileAsync` — every host below 1.10, and
  the floor is 1.2 — hands back the USER'S ENTIRE PRESENTATION. The merge
  removed the template block and nothing else, so on the file route it sent the
  whole deck back plus the clones: three rows into a forty-slide deck would have
  inserted forty-six. Keep the merged slides rather than remove the block, and
  the two routes become one case.
- **Blob downloads from the task pane are blocked in WebView2**
  (office-js#1511). Output goes into the deck, into a new presentation, or
  through `Office.context.ui.openBrowserWindow` — and the RUN RECORD goes on
  screen as selectable text between markers, which is the channel the probe
  already proved. A task pane also has no devtools a user can open, so anything
  a maintainer needs has to survive being copied out by hand.

## The lockstep rule (only item 1 is CI-enforced)

Any feature change updates, in the same PR:

1. **`docs/MANUAL.md`** — CI-enforced. `test/docs.test.ts` fails on a format or
   tag key the manual does not mention, on a button label the manual quotes that
   no longer exists in the pane, and on a manual that stops marking unbuilt
   sections as planned.
2. **`CHANGELOG.md`** — under `## [Unreleased]`, in the language a user would
   use, not the language of the diff.
3. **`docs/BACKLOG.md`** — an item that shipped is REMOVED, not ticked. Anything
   still listed is genuinely not done. A rejected idea moves to the rejected
   list with the reason, so nobody re-proposes it.
4. **`README.md`** feature table.

**Items 2 to 4 are on you, and this heading used to imply otherwise.** It read
"CI-enforced" over all four, and the only test touching `CHANGELOG.md` asserts
that the README links to it — nothing checks that a change added an entry. Six
pull requests went green on 2026-08-30 carrying four user-visible changes and no
changelog at all, and the green is precisely why nobody looked: a rule that
claims a machine is keeping it stops being a habit anybody keeps.

No test is proposed for it. Nothing in a diff distinguishes a feature change
from a refactor, and a check that guessed would be either noise or a rule people
learn to satisfy without meaning it. The honest fix is the heading.

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
- **A tool that edits the working tree will eventually commit into your work.**
  `scripts/mutate-core.mjs` wrote its mutants into the real source and restored
  them in a `finally`, which holds for a run that fails and not for one that is
  killed. On 2026-09-01 a run under `timeout` was interrupted, an unrelated
  `git add -A` swept the live mutant into a commit, and a dropped `.trim()`
  reached the branch inside a commit whose message was about a changelog entry.
  The script copies the tree now — but the general rule is the one that
  survives: `git add -A` commits whatever else is touching the tree, so stage
  paths when anything else is running, and check `git status` before a commit
  you did not build file by file.

  It cost nothing here only because the mutation run had itself reported that
  line as a SURVIVOR, so the gap it opened was already named.

- **Review the review.** Three rounds landed on 2026-09-01: six sweeps, then an
  adversarial review of those fixes, then a seventh sweep. The review found that
  two of the first round's fixes were INCOMPLETE — a crash still reachable
  because a magnitude bound cannot see an infinity, and a flag derived from a
  verdict string that hides the shape it was meant to detect — and the seventh
  sweep found a throw that killed a whole merge. A fix is not evidence that the
  defect is closed; the second pair of eyes is where that evidence comes from,
  and it is cheapest immediately after the work rather than a week later.

- **Record a diagnostic on BOTH populations, or it is not yet a measurement.**
  A value written only when something failed cannot be compared against
  anything. This repo shipped it twice: `withTimeout`'s call name reached only a
  `Timeout` message, so no successful call said what it was or how long it took
  and every number in `BUDGET` was a guess; and the row counts on a torn insert
  went out on the failure path alone until a test caught it. Ask the second
  question too — on how many of the runs that WORK is the field actually there?
- **Measure the artefact you hand over, not your intent.** `insertDeck` was told
  how many slides the plan BUILT where the package's own `sldIdLst` was the only
  honest answer, and on the file route the two disagreed by the size of the
  user's deck. The same rule caught the route defect above: what the code
  believes it produced and what it is actually sending are different numbers
  wherever anything can go wrong between them.
- **An instrument built before the failure that shapes it is built against a
  guess.** The debugging plan for this repo started as a port of a sibling
  project's apparatus — a 2000-entry ring with slice arithmetic and histograms,
  span subjects, a 445-line crash log — all of it sized to ITS failure history:
  276-entry runs, eight charts making per-chart host calls that all logged
  `index: 0`. A merge emits about ten entries and the per-row loop makes ZERO
  host calls, so most of it was cut and the 80% that closed the actual gap was
  fifteen lines at the `withTimeout` chokepoint. Take the sibling's RULES, which
  are free; take its code only where the same failure has happened here.
- **A probe's rule and a user's button need different rules.** `sweepPlan` capped
  its count when the deck had grown by more than the run added, reasoning the
  extras "are not ours". Sound for a probe sweeping seconds after it appends;
  wrong for an undo pressed after a coffee break, because the sweep removes the
  LAST slides and those now belong to whoever added the others. Before exposing
  a tested internal, re-ask its assumptions against the new caller's timescale.
- **Values never leave the pane.** Anything written to be copied out — the run
  record, a failure recipe — carries STRUCTURE only: column names and types, row
  counts, part names, block bounds. A mail merge's rows are salaries and
  customer lists, and this record is written to be pasted into an issue. Values
  do not change the parts, the relationships or the content types, which is
  where a rejected package goes wrong.
- **A gate that cannot fail is not a gate, and four here could not.** Found on
  2026-09-01 by asking of each one "what would I break to make this go red?" —
  which is the stash-and-re-run rule turned on the checkers themselves.

  - `test-count.mjs` read `numTotalTests`, which counts PENDING tests. All 23
    tests of `test/plan.test.ts` could be switched to `it.skip` with the count
    identical, the script exiting 0, the floor untouched and coverage nowhere
    near its thresholds. It counts tests that RAN now.
  - The docProps privacy guard asserts an empty list, and an empty list is also
    what a matcher that has stopped matching returns: `namesIn` could be
    replaced with `return []` and every case stayed green. The FILE LIST was
    already anchored against exactly this worry, which is what made the gap easy
    to miss — a guard can be defended against vacuity in one dimension and open
    in the other.
  - The no-Office-imports guard read prose-stripped source, and an import
    specifier is a STRING LITERAL, so `import "office-js"` in `src/core` could
    never match. Its own stripper's docstring says a guard that reads imports
    must read the raw source, and two other guards in the same file do.
  - `isMain` compared `import.meta.url` with `argv[1]`, and Node resolves the
    first through a symlink and not the second — so through `npm link` or a
    `.bin` shim every tool CLI printed nothing and exited 0. It had no test at
    all.

  The shape they share: each answers a narrower question than its name, and each
  answers it in a way that degrades to "fine". Ask what a broken input looks
  like, then feed it one.

- **`test-count` reads the working tree, so do not run it beside work in
  progress.** `--update` banked scratch files from a parallel agent into the
  floor, which CI then reads as tests having been deleted. Count against the
  committed tree — `git worktree add --detach /tmp/x HEAD` and run it there.

- **Held documents must not grow with the DECK either.** `Pkg.release` keeps the
  count flat against the record count and two tests pin that; nothing pinned it
  against the size of the template deck, and gathering the creation ids read
  every slide through the retaining cache. On the file route, where the template
  is the user's entire presentation, that is one parsed document per slide
  before a single record is merged. `Pkg.peek` reads without keeping. Both axes
  are pinned now — a gate whose name is wider than its coverage is the commonest
  defect this repo finds in itself.

- **Two passes over one file must agree about what the other has claimed.** The
  numeric pass refuses a chart value cell and the workbook's text pass ran
  afterwards and merged it — so the chart drew the template's figure while Edit
  Data showed something else, and closing Excel refreshed one from the other.
  Refusing is not the same as protecting: a pass that declines something has to
  say so to whatever runs next, or the decline is undone.

- **A fixture must prove it renders its own subject.** `scripts/pane-shots.mjs`
  had a state called `5-merge-done-undo` that drew no undo card, because the
  card needs a field no fixture set — so the one control that deletes slides was
  measured by nothing, under a name that says otherwise. Fixtures declare
  `shows` and `hides` now and the run fails when one stops being true. And the
  output directory is EMPTIED first: a PNG from an earlier build is
  indistinguishable from this run's, and one cost a false defect report.

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
5. Does `exportAsBase64Presentation` drop modern comments and
   `ppt/authors.xml` the way office-js#6867 reports the slide-level call
   doing? This add-in reads its TEMPLATE through that call and then clones
   what comes back, so a dropped part is a part every merged slide is
   missing. **Needs a deck somebody has actually commented on** — the arm
   compares against the same deck read through `getFileAsync`, and a deck
   with no comments answers NOT ASKED rather than giving it a clean bill.

Everything above is asked by the probe; `docs/PROBE.md` says how to run it and
`scripts/read-answers.mjs` reads the sheet. **Run it twice when an answer
decides something expensive** — one sheet is evidence about this host in that
minute, not about this host.
