# Changelog

Notable changes to SSF Merge. Newest first.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project uses [semantic versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
