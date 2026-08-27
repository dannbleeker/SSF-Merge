# Changelog

Notable changes to SSF Merge. Newest first.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project uses [semantic versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added — conditional slides, which the engine has always done

`prepare.ts` implemented conditional slides, `runPlan` reported
`unknownConditions`, `PaneState` carried `conditions`, and `main.ts` passed it
to both the preview and the merge. Nothing wrote it. The field was undefined in
every run that had ever happened, and the manual described the feature as
shipped — so a reader went looking for a control that was not there.

It is on **step 2**, under a line reading *Every slide, every row*: one dropdown
per slide in the block, offering the columns of the data attached. Step 2 rather
than step 1, because a condition names a COLUMN and only this step knows them.

A select, never free text. The engine matches a condition against a column name
exactly, so a typed name is a silent no-op discovered by counting slides in the
output. It does not remove the unknown-column case and is not meant to: a
condition is chosen from THIS paste's columns and the next paste may not have
them, which is said under the dropdowns before the merge and in the sentence
after it. That second half was missing too — `unknownConditions` reached the
pane from the day it was written and was read by nothing.

Conditions belong to the template, so a new paste keeps them; they are keyed by
slide number, so a changed block clears them. The test for that had to move to
an OVERLAPPING block before it could fail: with the old block 4-6 and a new one
7-9, a stale key of 5 is outside the block and the engine ignores it, so the
guard passed against the unfixed code and proved nothing.

`test/architecture.test.ts`'s known-unreachable list is empty again. It held one
entry, for exactly one change.

### Added — a merge that never finished can still be taken back

`undoInsert` is clamped against the deck's size BEFORE the run inserted
anything, and those numbers lived in a module variable. So a tab that died
during a merge left the user's deck holding 720 new slides and the pane with no
way to take them back — the slides were there, and the only thing missing was
two integers.

They go to `localStorage` before the insert now, and the next open offers the
run back. One key, one write. The sibling project's answer to the same shape is
445 lines because it is preserving a 300-entry narrative of a fifteen-minute
run; this is preserving `{deckAtStart, added, runId}`.

Three rules borrowed whole, all cheap: **never throw** (a store can be absent,
blocked by policy — a task pane is a third-party frame — or full, and none of
that is a reason for a merge to fail); **probe with a READ**, because a full
store answers no to "is there room" and yes to "is there a record", and the
record is the half that matters; and **validate on the way out**, since these
numbers authorise deleting part of a presentation.

### Added — a build stamp and an environment line

**PowerPoint caches the pane's HTML for ten minutes.** Open it too soon after a
deploy and a round tests code the host never fetched, with nothing saying so —
the pane looks identical and the log reads as a clean run of the wrong build.
The commit is now stamped in at build time and carried in the run record.

The run also opens with what the host IS: platform, version, **every**
PowerPointApi set it publishes, whether it clears the floor, which template
route it will take, and whether it can read the selection. Every set, not only
the ones we gate on — the gap between what a host has and what this add-in uses
is where the next unusable API is hiding.

Emitted **after** the run's mark, never at wiring time: the sibling's
environment line was written when its pane loaded and its run slice began
later, so it reached none of its archived rounds. Present in the code, absent
from every artefact anyone read.

### Fixed — a partial insert is reported in ROWS

`insertVerdict` grades slides, and for the probe that is right: it inserts a
two-slide fixture and the slides are the whole question. For a merge it is the
wrong unit and says almost nothing. **"719 of 720 slides landed"** means one
row's three-slide block became two — and every row after it still looks
correct, so the user finds the short one at slide 141 with no idea it was ever
going to be there.

It now reads **"2 of 3 rows landed complete; row 3 got 1 of its 2 slides. Take
the slides back and run it again."** Rows that got nothing are counted apart
from rows that got some, because a row with two of three slides is the worse of
the two — it looks finished.

Per-row slide counts come from the plan rather than being assumed uniform, since
a condition leaves a row shorter than its neighbours.

**The prefix assumption is named rather than hidden.** The count walks rows in
order and stops where the slides ran out, which is the reading if the host
truncated the package; nothing establishes that a partial insert truncates
rather than dropping a slide from the middle. It is stated and not measured
because it changes no advice — the answer is to undo and retry whichever row
tore — and where every row is the same size the count is position-independent
anyway.

**A refusal still reads as a refusal.** Zero slides landing is not "0 of 3 rows
complete"; that wording would bury the fact that the call was rejected.

### Checked — undo after a torn insert leaves no orphans

A review claimed a torn insert would strand "orphan slides from a half-landed
record" because `sweepPlan` clamps on counts. Checked rather than accepted: it
does not. `added` is the MEASURED deck delta, not what the plan hoped for, so
the sweep removes exactly what arrived, the partial row included. There is a
test saying so now.

### Added — the probe asks whether the template export drops parts

The one open risk the sibling sweep surfaced, now a question rather than a note.

This add-in reads its template through `exportAsBase64Presentation` and clones
what comes back. office-js#6867 reports the slide-level `exportAsBase64` omitting
modern comments and `ppt/authors.xml` — a sibling project triaged that as no
exposure and was right to, because it calls the API for a PICTURE of a slide.
Here a part the export drops is a part every merged slide is missing, silently,
in a file that opens cleanly. Nobody had checked whether the
presentation-level call behaves the same way.

The arm exports **every slide in the deck** and diffs its part list against the
same deck read through `getFileAsync`, which returns the package unfiltered.
Both cover the same slides, or the export would legitimately lack the parts of
slides it was not asked for. **Part names only, never content** — a part list is
structure, and the sheet is written to be pasted into an issue.

**Three states, and the third is the point.** A deck with no comments cannot
answer this, so `exportPartsVerdict` says NOT ASKED rather than reading the
absence as "the export keeps everything" — the same mistake as recording a
question the run never put as a `no`. Run it on a deck somebody has commented
on; `docs/PROBE.md` says so.

### Added — a merge can be taken back

`undoInsert` and `sweepPlan` were built, clamped and tested, and reachable from
nothing. The pane kept the numbers an undo is clamped against, `undoSummary` was
written and covered, and no view rendered either — so a real merge into
somebody's deck had no way back short of pressing Ctrl+Z 720 times.

A finished merge now offers **"Remove slides 13 to 732, which this merge
added"**, below the sentence saying what the merge did, so what happened is read
before the offer to undo it. A partial sweep KEEPS the offer, because slides are
still in the deck and the user is the only one who can finish the job.

### Fixed — a sweep could have deleted a stranger's slides

**This reverses a decision recorded in the tests, and the reversal is the
point.** `sweepPlan` capped its count at what the run added when the deck had
grown by more, on the reasoning that the extra slides "are not ours, so do not
take them". That reasoning inverts: the sweep removes the LAST slides, so if the
deck grew by ten and the run added three, the last three belong to whoever added
the other seven — it would delete their slides and leave the run's own in place.

Harmless while the only caller was a probe sweeping seconds after it appended,
where nothing has had time to arrive. Not harmless for an undo pressed after
looking through 720 slides, on a deck AutoSave has been writing to and a
co-author may have been editing. It refuses outright now, and the pane says so
rather than doing nothing quietly.

### Added — what was in the package, measured before it was sent

The package is handed to `insertSlidesFromBase64` and nothing keeps it, so when
PowerPoint answers `InvalidArgument` the file that caused it no longer exists
anywhere. Handing the bytes back is not available either — blob downloads from a
task pane do not work (office-js#1511).

So every insert now records what the package WAS while it still existed: slide
count, part count, and **byte size**, which is the number most likely to explain
an insert that died inside its budget and which nothing recorded. On both
populations, because a byte count with no successful run to compare it against
cannot say whether this one was unusual.

A refusal additionally records the **recipe** — block bounds, row count, column
names and types, condition count. The engine is pure and package-only, so those
rebuild the package offline in Node with no PowerPoint anywhere. **Structure
only, and that is a rule rather than an oversight:** a mail merge's rows are the
user's confidential data, and this record is written to be copied out of the
pane and pasted into an issue. Values do not change the parts, the relationships
or the content types, which is where a rejected package goes wrong.

### Added — a run record, and the merge finally says what it DID

**Every host call now names itself and its duration, on both populations.** Until
now `withTimeout`'s `what` reached exactly one place: the message of a `Timeout`
nobody sees unless the call failed. So no successful call in this codebase said
what it was or how long it took, and there was no baseline against which a
41-second insert is normal or alarming — every number in `BUDGET` was a guess
with no measurement behind it. A value recorded only on failures cannot be
compared against anything and is not yet a measurement.

Three states, not two: `answered`, `gave up waiting`, `raised`. A call that ran
out of budget and a call that threw are different facts about the host and want
different next steps.

**The line is written BEFORE the call and named for what it knows.** `issued`
is what it knows. A call that never answers is on the record while you are
still waiting for it — which is the whole point, and why the name cannot imply
completion.

**`src/core/trace.ts`** is the record: a capped array cleared at the start of
each run, one clock origin, payloads copied at write and at read, and a watcher
that cannot cost the log the entry it was writing. Deliberately small — a merge
emits about ten entries, and the ring arithmetic and histograms a sibling
project needs are sized to its 276-entry runs, not ours.

**The pane says what it is waiting on.** A merge is legitimately silent for up
to two and a half minutes (`BUDGET.file` 90s plus `BUDGET.insert` 60s), and a
frozen "Merging…" for that long is indistinguishable from a pane that has
wedged. It now names the call in flight.

**And the record is reachable.** A collapsed, selectable block holds the run log
between markers. A task pane is a nested cross-origin iframe with no devtools a
user can open and blob downloads from one are blocked (office-js#1511), so being
on screen to copy is the channel that works — the same one the host probe
already uses.

### Fixed — a merge that filled no placeholders reported success

`runPlan` has always returned `paragraphsMerged`, with a docstring saying "a
zero here on a real template means the fields never matched", and the merge seam
threw it away. So a merge over a template whose placeholders are spelled how its
author spelled them — the likeliest way a first run against a real deck goes
wrong — inserted every slide, changed nothing on any of them, and reported
"720 slides added".

The finished merge now reads **"720 slides added after slide 12 · no
placeholders were filled — check the spelling in your template"**. A zero is
said out loud rather than dropped as an empty clause: `0` is an answer, and here
it is the whole finding. Rows and slides skipped by a condition are named too,
so "8 rows" and "6 slides" reconcile instead of reading as loss.

### Fixed — an error message could put the whole merged deck on screen

Office echoes an argument back into `debugInfo`, and the argument to
`insertSlidesFromBase64` is the entire merged package as base64 — tens of
megabytes on a large merge. The path from there to the pane was `err.message` →
`state.notice` → a DOM text node, uncapped at every step, with the sentence
explaining the failure at the front and nothing after it readable. `readable`
now lives once in `src/host/errors.ts` and caps at 400 characters, counting what
it dropped rather than leaving a bare ellipsis.

### Fixed — a merge on an older PowerPoint would have duplicated the user's whole deck

`readTemplate` has two routes and only one of them was handled past the merge.
On `subset` — PowerPointApi 1.10 and up — `exportAsBase64Presentation` returns a
package holding only the template block, so removing that block leaves the
merged slides alone. On `file` — **every host below 1.10, and this add-in's
floor is 1.2** — `getFileAsync` returns the user's ENTIRE PRESENTATION, and the
merge removed only the template block from it.

So the package handed to `insertSlidesFromBase64` was the user's whole deck,
minus the template block, plus the merged slides. Three rows merged into a
forty-slide deck would have inserted **forty-six** slides: a second copy of
everything they had.

Not silent — the deck delta would not have matched and `insertVerdict` would
have said so — but by then the slides are in the deck, and "the merge duplicated
my presentation" is not a diagnosis anyone should have to make from a verdict
line.

The merged slides are kept now rather than the template block removed, computed
from the package's own `sldIdLst`, which makes both routes one case. And the
count handed to the host is the package's own, not the plan's: `insertVerdict`
grades the deck delta against it, so a number taken from anywhere but the
artefact makes the verdict a statement about the wrong thing — on the file route
the two disagreed by the size of the user's deck.

Found by reading the sibling project's debugging apparatus, not by a round.

### Added — the sibling sweep, and two findings it surfaced on its first run

`scripts/sibling-watch.mjs` reads PowerChart's curated tables — its triaged
office-js issues and the four host-answer tables that gate its fake against a
real PowerPoint — and reports anything with no row in `docs/SIBLING.md`. Weekly,
one issue reopened rather than duplicated. Raw file reads only, no GitHub API
and no token, and it never imports the sibling's code: a weekly job that
executes a file fetched over the network is a supply chain, not a sweep.

Seeded by running it against the live tables and triaging all **69** findings,
so the first Monday is quiet rather than a wall nobody reads. 52 are **no
exposure** — no shape work, no grouping, no rasterising, no pictures — and
saying so is the point: an untriaged finding is indistinguishable from an
unnoticed one.

**Two findings for us that the sibling had correctly marked no exposure for
itself:**

- **`exportAsBase64Presentation` may be dropping parts.** office-js#6867 reports
  that `Slide.exportAsBase64` omits modern comments and `ppt/authors.xml`.
  PowerChart calls that API for a picture of a slide, so it does not care. We
  call the presentation-level export to read the TEMPLATE WE THEN CLONE, so any
  part the export drops is a part every merged slide is missing — silently, in a
  file that opens cleanly. Recorded rather than fixed; it wants a probe question
  before the first real merge on a deck with comments.
- **The web uppercases tag keys** (office-js#6079) and then requires the
  uppercased spelling to read them back. Every key this engine writes is already
  uppercase, so this is safe by luck rather than by design — and now written
  down, because a lowercase key would go into the package fine and be
  unreadable.

A table renamed upstream throws by name rather than matching nothing: "nothing
new, every Monday, forever" is indistinguishable from a quiet week, and that is
the failure the whole apparatus exists to prevent.

`TRIAGED` is the source of truth and `docs/SIBLING.md` is downstream of it —
every reason opens with a verdict from a closed vocabulary, and the suite fails
when a finding we acted on has no line in the ledger. That check found thirteen
missing rows the first time it ran, against a ledger written the same morning.

### Added — a ledger for what we know from the sibling project, and a rule that keeps it honest

Most of what this repo knows about the PowerPoint host was learned by
PowerChart, a sibling add-in, over roughly 175 real-host rounds. SSF Merge has
run zero real merges. That debt was being paid by hand-copying facts, and
hand-copied facts go bad quietly.

**`docs/SIBLING.md`** is now the single place that knowledge lives: every
finding, where in PowerChart it came from, and what was done about it —
including the ones that are **no exposure** here, which is a real answer and the
one nobody writes down. An untriaged finding is indistinguishable from an
unnoticed one.

**The drift was measured, not assumed.** There were 44 citations across 24
files. Sorting them found two kinds: a single run's observation ("a by-id
clean-up reported 45 deletes and removed nothing") is true forever, and most
citations are that. A COUNT OF ROUNDS is a live counter. Four source comments
said "174 consecutive archived rounds" and one said "passed 174 of 174" — all
correct the morning they were written, all wrong the moment round 175 runs. One
had already drifted: the backlog paired a batch count from the measurement that
reported 29.2s with a maximum of 31.1s from a later round.

**The rule: a number copied from a live counter carries the date it was taken.**
Dated, it stops being a claim and becomes a recording — never false, only older,
and a reader can judge that. Where the number adds nothing, the phrasing that
cannot decay is preferred instead ("every rung answered in every archived
round"), and figures from two measurements are never paired.

`test/sibling.test.ts` holds it. Two things it had to learn the hard way, both
of which this repo already had a rule for: it is scoped to source comments and
skips its own file, because the files that STATE the rule quote the sentences
that break it — and it scans with comment prefixes stripped and lines joined,
because the claim that actually rotted wrapped across two lines and a
line-by-line version passed against the unfixed file.

### Removed — a Danish locale is no longer planned

Dropped at the owner's decision. The backlog entry that carried it was wrong
about the cost in the direction that matters: it claimed a string table already
existed, and none does. Every user-visible string in the pane is an inline
literal and a good number are assembled from fragments, so the work is building
the table rather than translating it — and the assembled sentences would not
survive extraction unchanged, because Danish pluralises and inflects
differently. It moves to the rejected list with that reasoning, so reviving it
starts from what it actually costs.

### Added — row filters

**A searchable checkbox list picks which rows merge.** Pasting 400 rows and
wanting 380 of them was previously a trip back to Excel to delete twenty and
paste again. The list is closed by default — a 340 px pane with 400 checkboxes
open is not a pane — and it says how many rows are in and how many are out on
its own header, so the state is readable without opening it.

The search box filters by any cell in the row, not only the first column, and
filtering does not change what is ticked: the ticks are the filter, the search
is a way to reach them. Untick a row while a search is narrowing the list and
that row stays unticked when the search is cleared.

**The list is capped at 60 rows shown and SAYS SO.** A pane that silently draws
the first 60 of 400 looks exactly like a pane that has 60 rows, so the
remainder is counted out loud under the list, with the suggestion to search.

**Every count downstream is the INCLUDED count.** The merge summary, the button
("Add 380 slides"), and the preview all read the filtered set, not the pasted
one. Unticking every row blocks the merge with a sentence saying so rather than
running a merge that adds nothing. A new paste clears the filter, because the
row numbers a filter holds are meaningless against different data.

### Added — two probe questions, and the deck read paged against the answer

**`deckSlideIds` pages the deck's id list, by position.**
`slides.load("items/id")` is the obvious way and it is the one office-js#4272
describes failing: past ~50 items the web host answers with FEWER than it has,
after a sync that succeeded. This add-in needed that list twice — to pick a
template block's ids, and to turn a selection into slide numbers.

What a short read costs here is not merely a smaller list. `blockIds` slices by
INDEX and `blockFromSelection` calls `indexOf`, so a short read that is not the
first n in deck order makes both answer the wrong SLIDE NUMBER — silently — and
the merge then clones slides nobody chose. A mail-merge template deck is exactly
the kind that gets large, so this is not a theoretical ceiling.

The ids come by `getItemAt` in pages of 20 now: a different code path from a
collection load, not subject to its limit, with `getCount` — a scalar, not a
load — as the authority on how many there are. **This did not wait for the probe
answer**, because paging is correct either way and the unpaged read is wrong if
the answer is bad.

**Two questions added to the probe.** `deckRead` asks whether a load answers
short at all, whether a short read is prefix-stable, whether it ever comes back
EMPTY (office-js#6363), and — so the sheet cannot be over-read — whether the
deck was even big enough to answer the first one. `insertWhileSelected` asks
whether an insert lands with a shape selected, and asks it WITHOUT selecting
anything: the workaround would be `setSelectedShapes`, the one call in this
family with a measured history of wedging the host.

The guard for that last property matched the snippet's own COMMENTS explaining
why it does not call `setSelectedShapes` — the fourth time a guard in this repo
has read prose as code, and the first since `without-prose.mjs` existed to
prevent it. The tool was there and I did not reach for it. It does now.

### Fixed — the delimiter sniff, and two things a review against the sibling project found

**The sniff sampled half a cell.** `parseDelimited` decided tab-versus-comma
from `src.slice(0, src.indexOf("\n") + 1)` — and that first newline can be
INSIDE the first header cell, because Excel writes a quoted newline whenever a
cell holds a line break and it is legal CSV. The sample then never reached the
tab, and a whole tab-separated table parsed as ONE column: every placeholder
unmatched, the merge button down, and nothing on screen saying why. It walks
with the same quote rule the parser uses now, which is the only sample that
cannot disagree with it.

A branch was written for the escaped-quote case and then removed when a revert
proved it could not fail: `""` toggles twice and nets to zero, which is the same
state the parser reaches by skipping both.

**The version floor was wrong, in the direction that turns users away.** It said
**1.3**, justified by `slide.tags` — and nothing in this add-in calls it. Merge
metadata goes into the PACKAGE, as `ppt/tags/tagN.xml` related from the slide,
which is the entire reason the engine never asks the host for something it can
put in the file. Every call `src/office` actually makes is **1.2**. Declaring a
floor higher than the truth excludes hosts that would have run the add-in
perfectly well, and does it by telling the user their PowerPoint is too old when
it is not. There is a test on the CLASS of error now, not just the instance: the
floor may not be justified by a call `src/office` does not contain.

**And the selection shortcut shipped unguarded.** `getSelectedSlides` is
PowerPointApi **1.5**, against a floor of 1.2. It went in on the strength of 174
rounds of evidence that it is not WEDGED — without anyone asking which version
introduced it. Safe to call and present are different questions. `canSelectSlides`
gates it now, the control is drawn only where the host has it, and
`selectedBlock` refuses with a sentence rather than a TypeError for a caller
that skipped the check. The two slide-number boxes work everywhere, which is
what makes the shortcut safe to be absent.

### Added — pick the template block by selecting slides

**Unblocked by evidence that already existed.** The backlog said this needed a
probe question before it needed code: `getSelectedSlides` is the obvious route,
and this host is documented to wedge its whole selection subsystem after
`setSelectedShapes` (office-js#3083, #3698). The question was whether the SLIDE
selection API is affected at all, or only the shape one.

A sibling add-in on the same host has been asking exactly that, every round, for
months. It runs a "selection ladder" — a read, `setSelectedSlides`, a read,
`setSelectedShapes([id])`, a read, `setSelectedShapes([])`, a read — and stops
at the first rung that goes silent. Across **174 consecutive archived rounds,
every rung answered**, in 550-710ms, with **zero refusals and zero silences**.
Its "edit the chart the user selected" scenario reads `getSelectedSlides`
exactly as this does and passed **174 of 174** — and it runs AFTER the ladder,
so the read survives even the call #3698 names.

So the wedge is not live on this host, `setSelectedSlides` was never implicated
in it even when it was, and this add-in never calls `setSelectedShapes` at all.
The measurement is cited at the call site rather than summarised here, because
the next person to doubt it should find the numbers next to the code.

**And the sibling had already paid for a gotcha this would have hit.**
office-js#2474: a `SlideRange`'s id is *not* the deck's id — it lacks the
`#XYZ` suffix — so `slides.getItem(rangeId)` answers InvalidArgument where the
deck's own id works. Closed `not planned`. The failure is the silent kind: the
id resolves to a null object and the slide is treated as gone. The ids happen to
round-trip on the web host today, which is exactly why it must not be left to
luck. `deckIdForSelectedSlide` matches a suffix-less id by prefix, and refuses
when two slides answer to one — guessing between them would name the wrong
slide, which is worse than refusing.

`blockFromSelection` turns the selection into slide NUMBERS, and refuses three
things rather than guessing: an empty selection, a slide the deck will not name
(dropping it would build a block out of whichever slides happened to resolve),
and a selection with a GAP. That last one matters most — a template block is
slides that repeat *together*, so closing the gap would add a slide to every one
of the user's rows that they never picked.

The control fills the two BOXES rather than committing a block, so the user
still presses "Use slides N to M" — the read that finds the placeholders.
Skipping it would leave the fields step listing nothing. It is a link beside the
boxes rather than the primary, because typing two numbers always works: if the
host ever does stop answering, the step still functions.

One thing is deliberately NOT claimed as tested. Clearing `state.block` when the
boxes are filled from a selection keeps that field meaning "a block whose
placeholders have been read" — but nothing observable distinguishes it from
committing the selection, since `chosenBlock` prefers the draft either way and
the template step's only way forward is the button that reads. It is stated in
a comment as defensive, and known to be, rather than guarded by a test that
would pass against both.

### Changed — one comment-stripper instead of three

Three guards in this repo have gone wrong the same way, each in a different
syntax, and each was found only when it went red on a file that was correct or
green on one that was not:

- `test/architecture.test.ts` forbade Office.js in the engine and matched the
  word "Office.js" in the paragraphs explaining WHY the engine avoids it — four
  correct files, red;
- `scripts/manifest-rules.mjs` forbade a `<Requirements>` block and matched the
  XML comment explaining why the manifest has none, so the generator refused to
  write a file that was exactly right;
- `test/release.test.ts` checked that the pre-flight runs before the tag is
  created and matched the YAML comment mentioning `gh release create`, so it
  compared the comment's position with the check's and reported the order
  backwards.

A file that explains itself is not a defect; a guard that cannot tell an
explanation from an instruction is. `scripts/without-prose.mjs` is the one
module now — `withoutXmlComments`, `withoutHashComments`, `withoutTsProse` —
and `test/without-prose.test.ts` drives each against the exact case that caught
its guard, so a stripper that stops handling its own case fails rather than
quietly putting a guard back where it started.

None of the three is a parser. Each is the smallest thing that makes its own
guard honest, which is why they are named for what they remove rather than for
a language.

Three private copies was three chances to write a fourth. The next person
reaching for one finds it.

### Added — a release workflow, and a gate on the thing every other gate misses

Every check in this repo reads the WORKING TREE. A user downloads the RELEASE,
and on a sibling project those two diverged twice — once shipping the DEV
manifests, and once with the README pointing at a `manifest-prod.xml` that was
not in the release at all. That second one stood for twelve days with a correct
release workflow sitting un-run, because nothing compared the documentation with
what was actually attached.

`scripts/release-assets.mjs` compares them, in both directions:

- an asset that is not a production manifest is refused outright — shipping a
  dev manifest points every installer at a localhost port nothing is listening
  on;
- every rule `scripts/manifest-rules.mjs` holds is applied to the FILE BEING
  SHIPPED rather than to the one in the tree, through an injected reader, so
  the workflow checks the bytes it is about to upload;
- and every production manifest the documentation tells a reader to download
  must be attached. The list of those is read out of the PROSE in
  `docs/MANUAL.md` and `README.md`, not written down a second time — a second
  list is a third thing that can disagree with the other two.

The workflow (`.github/workflows/release.yml`) is manual only, because a release
is a decision rather than a consequence of merging. It runs everything `test`
runs, regenerates the manifests and fails if the tree was stale, validates with
Microsoft's own tool on the exact file being shipped, runs the asset pre-flight,
and only then creates the tag — server-side through `gh release create`, since
the git proxy this project is developed through rejects a pushed tag. A tag made
before the checks is a release that has to be yanked instead of refused, and
there is a test on that ORDER.

Two bugs were caught in the workflow before it ever ran. `--generate-notes`
alongside `--notes-file` is ambiguous, and an indented heredoc keeps its leading
spaces — so the whole release note would have rendered as a grey code block. The
notes are written to a file and dedented now.

And the test for it read the workflow's own COMMENT rather than its steps: the
comment explains why the tag is created by `gh release create`, so "the
pre-flight runs first" compared the comment's position with the check's and
reported the order backwards. Third time in this repo that a guard has read
prose as code — after `manifest-rules.mjs` and `architecture.test.ts` — and the
same fix each time.

The manual now points at the latest release rather than at a raw file in the
repository.

### Added — the preview step, and a refusal of the design it was specified with

Step 3 shows one row on real slides, so all four steps of the pane are built.

**It merges.** Pressing "Preview the first row" runs the ordinary merge over a
one-row set and inserts the result at the end of the deck; "Remove the preview"
sweeps it away with the same clamped positional sweep an undo uses. What the
user looks at was produced by the code that will produce the other 239 slides,
which is the only thing that makes a preview worth having — a preview rendered
by some other route is a preview of something nobody is going to get.

**The specified design was refused, and the backlog and the rejected list said
so at the same time.** The backlog asked for "write one record's values onto the
real template slide, store what was there in a `SSF_MERGE_TEMPLATE` shape tag,
and put it back"; four items below, the rejected list forbids merging through
the PowerPoint API because setting a shape's text re-authors it
(office-js#5858 — custom bullets reverting to default discs, mixed runs
collapsing). Putting it back goes through the same API that did the damage, so
the TEXT would return and the FORMATTING would not — silently, on the master
copy every merged slide is cloned from, which is the one thing this product
exists to preserve. `TAG_TEMPLATE` is gone with the design that needed it, and
the refusal is written into the rejected list so it is not re-proposed.

Consequences of the design that ships, both stated on the screen rather than
discovered: the preview lands at the END of the deck, because that is the one
insertion point this add-in has tested on a real host; and a preview left behind
by a closed pane is just slides, so the card NAMES them ("Slides 13 to 15 are a
preview of the first row") rather than only saying one is showing.

Two labels changed with the design. "Put the template back" described restoring
something that is never taken, and is now "Remove the preview". The heading
"Preview is not built yet" is now "See one row before you commit".

**Making the preview real took the wizard's way forward with it.** Every other
step's primary advances; this one acts. Without a forward affordance step 3 was
a dead end, so the preview step carries a "Skip to the merge" link — rendered
only when the merge is actually reachable, so it cannot carry the user to a step
that refuses them.

The preview names itself while it runs — "Previewing…" and "Removing…" — like
the other two long host calls. Inserting a preview IS a real merge and can take
a minute on this host, and a button reading "Preview the first row", greyed out,
for the whole of it is the state a user cannot tell from a pane that has stopped
responding. Third instance of that gap; found by re-reading the diff rather than
by a failure.

`endPreview` checks the sweep's count rather than believing it. A sweep that
removed fewer slides than it asked for leaves part of the preview in the deck,
and the user is the only one who can finish that; the pane says so and stays in
the previewing state, which keeps the merge blocked.

### Added — the manifests, and an add-in somebody can install

Four files from one definition (`scripts/manifest-source.mjs`): the XML manifest
a person sideloads, and the unified JSON one an administrator deploys, each in a
development and a production flavour. Four hand-maintained files is four chances
for the production one — the only one anybody sideloads — to say something the
others do not, so `test/manifest.test.ts` fails when a committed file stops
matching what the source produces.

**No `<Requirements>` block, and that is the load-bearing decision.** The floor
is PowerPointApi 1.3 (`slide.tags`; everything else the add-in calls is 1.2, and
`getFileAsync` is a Common API PowerPointApi does not gate) and it is checked at
RUNTIME by `checkFloor`. A host that does not meet a declared requirement set
does not show the add-in at all — no ribbon entry, no error, nothing for the
user to report — where the runtime check names the missing version and what it
costs them. The rule refuses a requirement set in either format, and there is a
test that adds one back to prove it still bites.

`<Version>` is `1.0.0.0` and deliberately not the npm version, which is `0.0.0`.
Office rejects anything below 1.0 outright — "Manifest Version Too Low" — and a
sibling project shipped `0.1.0` in four manifests for the whole life of its repo
with a fully green suite, because nothing there had ever asked Microsoft.

**The `<Id>` GUID is pinned as a literal** in the source, in the rules and in the
test. A new GUID is a new add-in: every existing sideload orphaned, every user
removing the old entry by hand, nothing anywhere saying why.

Microsoft's `office-addin-manifest validate` runs in a CI job of its own, kept
out of `test` because it calls a Microsoft service and `test` is the check a
merge waits on. It cannot run in the development environment at all — the
service answers **403** through the egress proxy — which is exactly why
`scripts/manifest-rules.mjs` holds the handful of rules that can be checked
offline: a version below 1.0, a changed GUID, a declared requirement set, a
missing `ReadWriteDocument`, a production manifest pointing at localhost, and a
development manifest pointing at production.

The ribbon icons are **drawn in code** (`scripts/build-icons.mjs`) rather than
checked in as binaries: a 16, 32, 64 and 80 pixel PNG plus the monochrome
outline the unified manifest wants, from the pane's own two colours, through a
PNG encoder that is a raster, one zlib stream and three chunks. A binary in a
diff is a change nobody can review, so the test asserts the committed bytes are
exactly what the drawing produces and that the drawing is still the mark it
claims — one orange row and two pale copies of it, inset, at every size.

The first mark drawn was the pane's own tick, a single bar across the middle,
and it read as a **minus sign**. Three rows say what the product does.

Also: `vite.config.ts` pins the dev server to port 3000, because that is the
origin the development manifest names and a manifest pointing at a port nothing
serves is a blank pane with a generic error. `.prettierignore` gives the
generator sole ownership of the JSON manifests — Prettier collapses their short
arrays, so with both tools owning the file `npm run manifests` and
`npm run format` leave the suite red whichever order they run in. And ESLint's
default-project file cap is raised: four new scripts pushed `scripts/` past
eight, at which point every file over the line reports the linter's own limit as
though it were a defect in the code.

### Fixed — what an adversarial review of the commit above found

Five lenses over the diff, each finding verified by three independent skeptics;
sixteen findings survived a majority, three of them blocking. The pane could not
be driven by the suite at all, so `test/pane-wiring.test.ts` now drives the real
`main.ts` with only the two Office-touching modules mocked, and
`test/office-merge.test.ts` covers the refusal paths `src/office` answers with.

**A merge that raised left the pane dead.** `merge()` awaited `runMerge` with no
catch, so a rejection skipped the `draw()` below it and the button kept the
hand-written `disabled = true` / "Merging…" forever, with nothing shown to the
user and `last` — the numbers a positional undo is clamped against — never set.
The path was not exotic: the commit above made `readTemplate` THROW its refusal,
and `runMerge` awaited it bare while `inspectBlock` next to it wrapped the
identical call. Both now answer instead of raising, and `merge()` catches as a
backstop; because a raise on this host does not mean nothing happened, it counts
the deck again and keeps whatever landed, saying so.

**The merge button came back mid-merge.** The only thing stopping a second press
was a `disabled` flag on a DOM node every later `draw()` replaced with one
`primary()` had re-enabled — so going Back and Continue during a two-minute
merge handed the user a live "Add 720 slides" over a run in flight, and their
deck got both inserts with one set of undo clamps. The run is in the STATE now
(`running`), so `primary()` can see it. A completed run disarms the button too
(`added`): redrawing a live "Add 720 slides" beside a notice saying 720 were
added is how a deck gets 1440.

**The architecture guard passed against the exact defect it names.** It sliced
`readTemplate`'s body to the next `/**`, which is where the next DOCBLOCK starts
and not where a function ends. With `readTemplate` reverted to the counting loop
and a helper below it holding the strings the guard greps for, all seven
assertions passed. Slicing to the next top-level `export` was tried and is also
wrong — a plain `function` is not a stop. It brace-matches now, over a copy with
strings and comments masked, skipping the parameter list (the first `{` after
the signature was `block: { from: number; to: number }`, which made the guard go
red on a correct file — the failure mode this file already has a paragraph
about).

**The caret jumped to the end of the box on every keystroke.** `render` builds
fresh elements, so each draw destroyed the focused control; the first version
focused the replacement and sent the caret to `value.length`. Typing 5 into
"4|6" gave 456 and the next digit landed after the 6, and an edit in the middle
of a pasted table scattered the rest of the line to the end — where
`readPastedTable` then parsed the corrupted text into the records the merge
runs on. `draw()` carries the selection across every redraw now, including ones
the user did not cause: the deck count resolves a second or two after the pane
opens, and its redraw was swallowing the next digit typed. The two boxes are
`type="text" inputmode="numeric"` because `type="number"` answers `null` for
`selectionStart` and throws on `setSelectionRange`.

**The block was committed after the boxes had moved on.** `useBlock` captured
the block before a read that takes seconds and wrote it back afterwards, with
nothing disabled meanwhile — so retyping mid-read left `state.fields` describing
one block while `chosenBlock` answered another, and the merge runs on
`chosenBlock`. Input is refused while a read is out, the answer is discarded if
the boxes changed anyway, and the flag is cleared in a `finally`.

**`inspectBlock` caught one call of three.** `Pkg.open` and `prepareBlock` were
awaited past its `try`, and both raise — on bytes JSZip cannot open, or a zip
with no `ppt/presentation.xml`. Those rejections reached `void useBlock()` with
no handler, leaving the pane reading "Reading the slides…" for the rest of the
session.

**The deck bound refused blocks that exist.** `deckSize` was counted once at
pane load and never again, so a user who added slides and came back was told
their block ran past the end of a deck that no longer had that size, with no way
to correct it short of reopening the pane. It is a WARNING now, not a refusal —
the authoritative check is `blockIds` against ids the host listed a moment later
— and the count is re-read on every press of "Use slides N to M".

**The paste box had no label at all.** A `div` with a sibling caption, where
`blockControl` two functions up correctly used a `label`, so a screen reader
announced "edit, multiline, blank" and clicking the caption did nothing.

Three tests were decoration and are not any more. `sets the box's VALUE` passed
against `setAttribute` too, because a fresh input reflects the attribute into the
property. Two of `readBlockDraft`'s refusals asserted only that the MESSAGE
contained a digit — satisfied by the literal 1 in "numbered from 1" — so an
implementation returning a block *and* a complaint passed the suite. And the
orange budget counted elements rather than holders, which meant a template
missing two columns read as two oranges: an ordinary state that no fixture in
the test or the screenshot script happened to reach, so CI reported neither the
violation nor the miscount. It counts holders now, which is what `orangeHolder`
already models, and two-unmatched is in the sweep.

Every refusal sentence names what the user typed. "Slides are numbered from 1."
is a true sentence that says nothing about the boxes in front of them, and the
manual promised numbers for all four cases while two carried none.

### Fixed — the merge run built its own slide ids, and PowerPoint would have refused every one

`runMerge` turned the template block into ids by counting —
`for (let n = from; n <= to; n++) ids.push(String(n))` — and handed
`["4", "5", "6"]` to `SlideCollection.exportAsBase64Presentation`, whose typings
say it *throws an `InvalidArgument` exception if provided slide IDs or `Slide`
objects are not found in this collection*. A slide id on this host looks like
`256#3561048925`. `tsc` could not see it: both sides are `string`.

All three answer sheets under `docs/host-answers/` report PowerPointApi up to
1.10, so `chooseDeckSource` returns `subset` on the owner's own host and the
first press of the merge button would have thrown.

- `readTemplate` takes slide NUMBERS now, so no caller can pass ids at all —
  the signature is the guard, and `tsc` rejects the old call outright.
- Its subset branch asks the host, loading `items/id` and passing what came
  back. `blockIds` in `src/host` is where that is checked: a block running past
  what the host listed, or an id the host would not name, is a sentence rather
  than an `InvalidArgument` from inside a callback.
- `test/architecture.test.ts` holds the shape in source, since `src/office`
  cannot run in the suite.

### Added — the pane's controls

The two things in front of the merge button. Steps 1, 2 and 4 are now reachable
from the screen.

- **Step 1 — the block.** Two slide-number boxes, read by `readBlockDraft` and
  checked as they are typed: a block that ends before it starts, a slide 0, a
  fraction, or one running past the end of the deck is refused in a sentence
  naming both numbers. Nothing is said while a box is still EMPTY — the boxes
  are filled one at a time, so a half-typed entry is not a mistake, and a form
  that turns red on the first keystroke is wrong more often than its user is.
  The draft is held as STRINGS, because `""`, `"-"` and `"0"` are all states a
  box passes through and a number type has nowhere to put them.
- **Step 2 — the data.** A paste box, read by `readPastedTable`. ONE parse, so
  the columns listed, the row count, the number on the button and the records
  the merge runs on cannot disagree. What it shows under the box is the COLUMNS,
  not just a count: a copy that came through as plain text parses into one
  column and a row count alone looks healthy when that happens. A header row
  with nothing under it is refused rather than counted as zero rows.
- **The fields step stopped guessing.** `inspectBlock` does the same read and
  the same `prepareBlock` the merge does and stops before the plan, so pressing
  "Use slides N to M" lists the placeholders actually in those slides. One
  template read per press, not per keystroke — which is why the block is
  committed on the button rather than as the boxes are typed.
- **A link back on every screen but the first**, rendered before the primary so
  the primary stays the last thing on the screen.
- `nextStep` bounds the walk. `advance` was `order[order.indexOf(from) + 1]`,
  and `indexOf` answers -1 for anything that is not a step — so a stray
  `data-action` sent the user to step 1 with their block and their data still in
  state, a wizard that resets itself and looks like it lost the lot.

Two labels were promises the code does not keep, and both were changed rather
than left to be pressed. The preview step's button said "Preview the first row"
and its heading "See one row on the slide" while nothing writes to a slide; the
screen now says the preview is not built, in the heading and the body, and its
button carries the user on to the merge. The fields step's button said "Attach
data" after the data was attached, which reads as a step that did not take.

An input's `value` is set as the PROPERTY, never the attribute — a `textarea`
has no value attribute at all, so one helper serving both controls has to write
the property. (This paragraph first gave a different reason, that the box would
"snap back" to the attribute's value; it would not. `render` empties the root
and builds fresh elements, whose dirty-value flag is unset, so `setAttribute`
would have worked for the inputs. The test named for the rule passed against
the wrong implementation for the same reason, and now asserts the attribute is
absent.)

`scripts/pane-shots.mjs` renders eleven states now — including the two the
controls added and the one where the host refused — and the orange-budget test
sweeps them too. A budget checked over the states that existed when it was
written stops covering the pane the first time the pane grows.

### Added — the merge run

The seam where the pane, the host layer and the engine meet. `runMerge` counts
the deck first (undo is clamped against that number), reads the template's bytes,
does the whole merge inside the package where nothing can be refused, hands
PowerPoint one deck in one call anchored after the last slide, and reads the
DELTA rather than the absence of an error.

- `prepareBlock` turns slide NUMBERS — the only numbering a user can see — into
  package paths, and refuses in sentences the pane shows as they stand: a block
  that runs off the end, one that ends before it starts, slide 0, and a block
  with no placeholders at all, which the engine would clone happily into N
  identical copies.
- `Pkg.removeSlide` takes the template slides out of the produced package.
  Inserted whole it would put the user's own placeholder slides back into their
  deck after every run. All five references go — the id list entry, the
  presentation relationship, the content-type override, the slide's own
  relationships and the part — plus the notes page, which belongs to exactly one
  slide.

  The alternative was to insert everything and name only the copies through
  `insertSlidesFromBase64`'s `sourceSlideIds`. That needs ids in the host's own
  `256#3561048925` spelling CONSTRUCTED for a package not yet in the
  presentation — an assumption no round in a real host has tested, whose failure
  mode is `SlideNotFound` and nothing inserted.
- The pane's merge button calls it. The controls that set the block and attach
  the data are the next increment, so the button cannot enable from the screen
  yet; the manual says so rather than implying otherwise.

Writing the `removeSlide` guard found the guard itself incomplete: with the
id-list line removed every assertion still passed, because `slidePaths` resolves
each relationship and SKIPS the ones that answer nothing — so a dangling
`<p:sldId>` reads as a tidy deck while `presentation.xml` references a
relationship that is gone, which PowerPoint refuses. The test walks the id list
directly now.

The architecture guard widened with it. It required every `src/office` file to
import from `src/host` and refused the merge run for taking its decisions from
`src/core` instead — the guard being narrower than its own reason, not the file
being wrong.

### Fixed — a merge held every slide it produced

The last of the bug hunt. `Pkg`'s document cache is also its dirty-part set, so
nothing ever left it: a run kept one live xmldom Document per output slide on top
of the zip's own copy of the same bytes. Measured on a 300-paragraph slide, 300
clones held **440 MB** of heap and 400 clones **591 MB**, against **54 MB** either
way once released — flat rather than growing with the record count, which is the
property that matters inside a task-pane WebView.

`Pkg.release` writes a part back and drops it; `runPlan` releases each slide, its
relationships, its notes page and the notes page's own relationships once nothing
will read them again. Parts the run keeps amending — `[Content_Types].xml`,
`ppt/presentation.xml` — are deliberately not released.

The guard counts held parts rather than heap, because a memory assertion would be
flaky and what is actually claimed is that the count does not track the record
count. **It caught the first fix being incomplete**: releasing the slide, its rels
and the notes page still left one document per record, because `cloneNotesSlide`
also edits the notes page's `.rels`.

Two numbers in the hunt's report are not repeated here because this repo could not
reproduce them: it claimed 17-18x and an out-of-memory kill at 400 records under a
2 GB cap. Measured here it is 8x at 300 and 11x at 400, and 400 completed. The
direction was right and the magnitude was not, so the magnitude measured here is
what is written down.

### Fixed — data, text and the pane

The rest of the bug hunt's confirmed findings. Every one produced silently wrong
output rather than a failure anybody would see.

- **A placeholder whose name is not English never merged.** The field pattern
  was `[\w.]`, which is ASCII-only and stays ASCII-only under the `u` flag, so
  `{{Beløb}}` and `{{Måned}}` were invisible: `fieldsIn` never reported them,
  the pane could not flag them as unmatched, and the literal braces printed on
  every merged slide. On a product whose first users write Danish that is most
  of a template.
- **An impossible date was rolled into a real one.** `Date.UTC` normalises
  rather than rejecting, so `29/02/2025` rendered as "1 Mar 2025" and
  `31/04/2026` as "1 May 2026" — dates a reader believes. The components are
  read back now and the cell is returned untouched unless it survived.
- **Every named or ISO date was a day early east of UTC.** They were parsed in
  the local zone and printed from UTC fields, so in Europe/Copenhagen — this
  project's own locale — `1 Mar 2026` rendered as `28 Feb 2026`. CI runs in UTC
  and the only date assertion used the form the spec parses as UTC, so nothing
  caught it.
- **`numericValue` could not read a number with more than one thousands group.**
  `replace` without `/g` changed only the first separator, so `1,234,567` became
  NaN while `detectType` still called the column a number — half a column
  rendered formatted and half rendered raw.
- **A format asking for impossible decimals killed the merge.** `number:-1` is
  natural to write (Excel's ROUND takes negative digits) and `toFixed` throws
  outside 0..100, on a path whose own contract is to return the value unchanged.
- **An invented column name could steal one a real header owns.**
  `["Name", "Name", "Name 2"]` produced `Name, "Name 2", "Name 2 2"`, so a
  template's `{{Name 2}}` bound to the second `Name` and printed the wrong
  column on every slide.
- **`undoInsert` let a timed-out delete escape**, skipping the confirming
  re-count that `insertDeck` thirty lines above always did. The caller was told
  the undo failed while the user's slides were already gone, with no count of
  what went — on a host `CLAUDE.md` records as answering late on work it had
  actually done.
- **The pane said "1 placeholders".** The zero case was special-cased and the
  one case was not, and the screenshot script only ever renders three.

One guard written for the undo fix was DISCARDED as decoration: it re-asserted
`sweepPlan`, which is already covered, and would have passed against the unfixed
file. `src/office` cannot run in the suite, so the rule is held by a source scan
instead — every host call that changes the deck sits inside a `try` and counts
the deck again afterwards.

### Fixed — the package layer

A bug hunt across the engine found six defects here, every one reproduced before
it was fixed and none visible from the engine's own output: the merge looked
like it had worked.

- **`element()` searched DESCENDANTS where callers meant children**, which is
  one root cause behind two shipped defects. A slide's `<p:cSld>` contains the
  whole shape tree, so `element(cSld, "tags")` found a SHAPE's tag reference and
  wrote the slide's merge metadata into that shape's part — leaving the slide
  with no slide-level tags, which is the exact read undo depends on — and
  `element(cSld, "extLst")` appended the creation id inside `<p:spTree>`, where
  PowerPoint does not look for one. `child()` and `children()` are the fix.
- **A clone inherited the template's tag relationship.** The `.rels` are copied
  verbatim, so every merged slide wrote into the TEMPLATE's `tagN.xml`: all but
  the last record's tags were overwritten, and the user's own template was
  stamped as merge output and matched by undo. A copy starts with no tags now.
- **`setCreationId` emitted `xmlns:p14` twice**, which XML forbids outright, so
  PowerPoint rejected the whole package without saying which part was wrong.
  `createElementNS` already binds the prefix.
- **`mergeTagPart` re-escaped the foreign tags it kept**, turning `Ben & Jerry`
  into `Ben &amp; Jerry` on screen after one merge and `Ben &amp;amp; Jerry`
  after two — and it insisted on one attribute order and a self-closing tag, so
  PowerPoint's own legal spellings matched nothing and the foreign tag was
  silently DROPPED. It parses now instead of pattern-matching.
- **Notes pages were never merged.** A copy gets its own notes slide precisely
  so the copies can differ; unmerged, a template whose notes read
  `Call {{Name}} afterwards` shipped that verbatim on every merged slide, in the
  presenter view and on every printed handout.

One guard among these was written and found to be DECORATION: the shape-tags
test invented a relationship id that resolved to nothing, so the old code fell
through to the right answer by accident and the test passed against the unfixed
file. It reproduces properly now.

### Added

- The task pane: four steps, the SSF visual system, English, at the width a task
  pane is actually dragged between. The step machine, every label and the
  renderer are checked in the suite; `main.ts` is the only file that touches
  Office.js and `test/architecture.test.ts` holds that.
- `scripts/pane-shots.mjs` renders every pane state at 320 and 512 px. The pane
  is the one surface the suite cannot judge — jsdom has no layout and no colour
  — and its first run found a defect no test had.
- `npm run dev` and `npm run build`, and the Pages workflow now builds the pane
  instead of copying a placeholder. A pane change ships through main with no
  re-sideload; only a manifest change costs that.

- The host layer, split the way the probe was: `src/host` holds the decisions as
  pure functions the suite checks, `src/office` holds the Office.js calls and
  decides nothing. `test/architecture.test.ts` holds both directions — an
  Office.js import in `src/host` makes a rule untestable, and a rule
  reimplemented inline in `src/office` looks tidier and rots quietly.
- `capability.ts`: the version floor and where the template's bytes come from.
  The floor is **1.3**, read off the calls the add-in makes rather than picked,
  and checked at runtime rather than declared in the manifest — a declared
  requirement set the host does not meet makes the add-in vanish from the ribbon
  with no diagnostic at all.
- `timeout.ts`: a budget per call rather than one number. An insert gets sixty
  seconds because thirty was measured too short; a count gets fifteen.
- `powerpoint.ts`: reading the template, inserting after a `targetSlideId`, and
  positional undo that counts the deck again afterwards rather than believing a
  delete that raised nothing.

### Changed

- `sweepPlan` moved from `verdicts.ts` to `undo.ts`. It is the real undo's rule
  as much as the probe's, and a file about what a probe answer MEANS is the
  wrong home for the one that authorises deleting somebody's slides. The test
  count was pinned across the move: 132 before, 132 after.
- `docs/BACKLOG.md` said the requirement floor was 1.4, that image fields were
  blocked on probe question 4, and nothing about how a per-call budget differs
  from waiting longer on a dead call. All three corrected.

### Fixed

- The architecture guard matched Office.js inside STRING LITERALS, so a verdict
  whose text names office-js#6105 failed a file with no imports at all. Literals
  are stripped now, as comments already were. The import check reads the raw
  source instead, because an import specifier is itself a string.

- `insertVerdict` read the presence of an error as decisive and ignored the deck
  delta, so an insert that timed out having landed both its slides was reported
  as a refusal. One misread arm then produced three false statements in the same
  run: that our package had been refused, that the collision arm disagreed with
  the fresh one, and that the theme was the difference. Its own docstring
  already said the delta is the evidence; the code agrees with it now. A raise
  that landed nothing, or only part of what was asked, is still a throw, and the
  partial count is reported rather than hidden.
- The probe's insert budget was 30 seconds, which expired once on a call that
  worked. It is 60 now: a budget that fires on a successful call produces a
  false refusal, which is the more expensive direction.

- Questions three and four could not be reached: a shape proxy does not survive
  a `context.sync()` on PowerPoint for the web, which answers `5010:
  InvalidParam passed to GetItem(id)` the first time a shape created a sync
  earlier is touched again. Both experiments are queued in the batch that
  creates their shape now, which is also the shape a real merge has. They draw
  on a slide the probe inserted, so the positional sweep removes them, and they
  do not draw at all when no insert landed.
- The reader scored question four against `"Hello Ada here and 1-2".replace("2",
  "BBB")`, a string neither offset model produces. Both experiments' text and
  both predictions are single-sourced now, and a guard checks each prediction is
  what applying that model by hand actually gives.

- The probe reported a verdict on a question it had not asked. The tag read
  lands on the last slide in the deck, which is the inserted one only if the
  insert worked; the first real sheet had every insert throw, so that read fell
  on the user's own title slide and the reader announced that the metadata
  scheme needed rethinking. The scheme had never been tested. `tagVerdict` now
  answers "not asked" whenever the slide carrying the tag did not land.
- The probe's fixture deck carried `<a:themeElements/>`, which is
  schema-invalid: `CT_BaseStyles` requires clrScheme, fontScheme and fmtScheme,
  all three mandatory, and `KeepSourceFormatting` is precisely the path that has
  to import the source theme. Checked against a deck PowerPoint itself accepts.
  The fixture also gained the docProps parts and a real `xfrm` on its text box,
  so a rejection has fewer places to hide.
- Four separate host calls in the substring probe shared one `catch`, so a throw
  named none of them and the sheet said `InvalidArgument` about a statement
  nobody could identify. Each call now stamps what it is doing, and the error
  carries it.

### Added

- A control arm on question one, and the rule that reads it. Inserting the
  presentation's own bytes asks a question only one of the two readings of
  `InvalidArgument` survives: a deck PowerPoint saved seconds ago cannot be a
  malformed package, so a host that refuses it is refusing insertion itself.
  `insertionBlame` says OURS, THE HOST, or CANNOT TELL, and never guesses.
- A second insert of the same package under `UseDestinationTheme`, which does
  not import the source theme where `KeepSourceFormatting` must.
- The generated snippet is typechecked against the real Office.js types in CI.
  It sits outside tsconfig's include and is pasted into an editor that runs it
  before anyone reads it, so nothing else would catch a misspelled option key.

- The host probe: a Script Lab snippet that asks the four questions only a real
  PowerPoint can answer, and a reader that interprets the sheet it produces.
  The snippet collects observations and makes no judgements; every reading
  happens in the repository, where it is tested.
- The merge plan: a block of contiguous slides repeated once per record,
  record-major so a record's slides stay together, with conditional slides
  skipped in place and a policy for what an empty cell means.
- The runner that carries a plan out against a package: clone the template
  slide, merge the copy, tag the copy. The template is never touched, so it can
  be merged again.
- The package layer: reading and writing a .pptx as parts, relationships,
  content types and the slide id list, with base64 in and out.
- Slide cloning, with its own relationships, content type, presentation
  relationship, slide id, notes page and a **fresh `p14:creationId`** per copy.
- Tags written into the file as `ppt/tags/tagN.xml`, referenced from the slide's
  `custDataLst`, merging with tags another tool already wrote.
- Run-aware placeholder replacement that survives PowerPoint splitting
  `{{FirstName}}` across runs, and leaves every run's properties untouched.
- Table parsing with column type detection, and formatting for numbers, dates
  and case. An ambiguous slash date is refused rather than guessed.
- The documentation set: manual, backlog, changelog, and a test that keeps them
  in step with the code.
- Hosting on GitHub Pages at `ssf-merge.struktureretsundfornuft.dk`.
- A manual trigger on the CI workflow, so a commit whose push event was lost can
  still be tested.

- Prettier and type-aware ESLint, both run in CI. `no-floating-promises` is on,
  because the engine is async throughout and a missed `await` there merges the
  wrong thing quietly rather than throwing.
- Coverage floors on `src/core`, and a floor under the number of tests so a
  reorganisation cannot delete cases silently.
- Dependabot, with a log in `docs/DEPENDENCY-ALERTS.md` where every alert gets a
  written reading, "no exposure" included.
- Contributing guide, security policy, pull request template and code owners.

### Changed

- Vitest moved to 4.x, so the repository does not start two majors behind.

### Fixed

- `Pkg.text` and `Pkg.copyPart` read edits that had not yet been written back to
  the zip. Three tests were passing on the version from disk.
- Five unnecessary type assertions in the package and XML layers, found by the
  new linter.
- The no-Office-imports guard matched the words "Office.js" and "PowerPoint.run"
  in the comments explaining why the engine avoids them, and failed on four
  correct files.
