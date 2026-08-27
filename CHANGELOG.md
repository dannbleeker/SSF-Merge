# Changelog

Notable changes to SSF Merge. Newest first.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project uses [semantic versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
