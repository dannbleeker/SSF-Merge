# Changelog

Notable changes to SSF Merge. Newest first.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project uses [semantic versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed — the deploy no longer races the tests

Pages ran on every push to `main` with nothing between the push and the live
add-in. CI ran too, but CONCURRENTLY: a commit could be serving from the
production origin before its tests had finished, and if they then failed, the
broken pane was already what PowerPoint loaded.

Checked rather than assumed — every commit on `main` has been green on both, so
this has never happened. It is the ORDERING that was missing, not the tests.

The deploy job now waits on a gate job running the same five checks CI runs.

They are listed in both workflows rather than wrapped in one `npm run gate`,
which was the first attempt: a script chaining `npm run a && npm run b` cannot
run on the maintainer's own machine, because npm spawns a shell for the `&&` and
AppLocker refuses it. A gate that only ever speaks through CI is worse than a
duplicated list. `test/release.test.ts` holds the two lists against each other —
same commands, same order — so they cannot drift, and asserts the deploy still
declares `needs: gate`.

### Added — a new source directory cannot go unmeasured in silence

`src/core/trace.ts` names this hazard in its own docstring, and says it chose
where to live because of it:

> the coverage config's `include` is a fixed list of three globs — a new
> top-level directory would be measured by nothing, and an uncounted module is
> how a threshold quietly stops meaning anything.

Nothing enforced it. Four directories exist and all four are accounted for —
three measured, `src/office` deliberately not — so this costs nothing today. The
day somebody adds a fifth, the four thresholds go on passing while saying
nothing whatever about it.

The list lives in `scripts/coverage-scope.mjs`, read by the config and by the
guard, with the reason for each exclusion beside it — a glob cannot carry one.
So adding a directory is a decision somebody writes down rather than a line they
forget.

Importing `vitest.config.ts` into the test was the first attempt and does not
work: pulling it into a test pulls it into the TypeScript project service, and
lint refuses a file that is in two projects at once. Sharing the list is the
better answer anyway — one definition, and the guard still measures it against
what is actually on disk.

It holds in both directions: a directory measured by nothing fails it, and so
does an exclusion that has stopped being true — a reason for not measuring a
directory that no longer exists, or that is measured after all, reads as a
considered decision and is a leftover.

### Added — a guard on the guards

`without-prose.mjs` is the shared stripper behind ten architecture guards, and
the module says plainly that none of its functions is a parser: each is the
smallest thing that makes its own guard honest. That trade has a boundary, and
it was not written down.

`withoutTsComments` finds block comments with a regex, so a `/*` inside a STRING
opens one and the next `*/` closes it. Everything between them leaves the file
the guard is reading, and the guard then reports that the file does not do a
thing it plainly does. That is the same failure direction three separate guards
in this repo have already taken, and the one that looks like success.

**No file in this repo trips it**, and a sweep now says so rather than leaving
it to be assumed: every declaration a file makes outside a comment has to
survive the stripper. It fails the day one does, naming the file and what was
lost.

The boundary itself is pinned by two unit tests — the case that breaks it, and
the near-miss that does not, since it takes an opener AND a closer, which is why
this has never bitten.

The sweep allows a declaration QUOTED IN A COMMENT to disappear with the
comment, because this repo's docstrings quote them constantly. Without that it
reports `scripts/sibling-watch.mjs`, whose docstring quotes
`export const NAME` — and a sweep that cries wolf is one somebody widens until
it is quiet.

### Added — the release gate refuses a production manifest served over http

Office fetches every address a manifest names and requires HTTPS for all of
them. A production manifest on `http://` fails Microsoft's validator, and
sideloaded anyway it fails the way this rule set's other entries describe: no
ribbon entry, no error, nothing to report.

`checkManifest` matched `https?://` already — the pattern was written wide — but
the only rule reading it was the localhost one, so an insecure production
address passed with nothing to say about it. `PROD_ORIGIN` is one constant in
`manifest-source.mjs`, which is this file's own test for whether a rule earns its
place: one edit away from being shipped.

The namespace declarations are stripped before the addresses are read, and that
is the whole difficulty of the rule.
`xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"` is an IDENTIFIER
— never fetched, http by definition, and not ours to change. A rule that read
those would fire on every manifest ever written and would have been deleted
rather than fixed, so a test asserts the real manifest still passes WHILE
carrying one.

All four generated manifests pass unchanged.

### Fixed — the picture picker had no name

Every control on this pane sits inside a `<label>` carrying its caption. The
picture picker does not: its block is a `<div>`, because it holds a tally and a
missing-file list as well, so the caption sat beside the input and was attached
to nothing underneath it.

A file input with no accessible name is announced as an unlabelled button — the
user is told there is something there and not what it does — and it is the only
control on the pane that attaches the pictures.

Named with an attribute rather than by moving the caption into a `<label>`,
because the caption is one of several children and re-nesting it is a layout
change a test cannot verify.

Found by sweeping every step for three things a browser or a screen reader will
act on: a name for every control, no duplicated id, and no button that says
nothing. One fault across all five steps, and the sweep is now a test — the
control that was missing had been added later than the pattern, which is the
kind nobody writes a case for.

### Fixed — a failed merge could explain itself as "[object Object]"

`readable` turns whatever was thrown into the sentence the user is given. Its
fallback was `String(e)`, which for an object reaches Object's default
stringification.

So a thrown `{ message: "InvalidArgument", code: 5 }` arrived as
**"[object Object]"** — discarding the message sitting inside it — and a raise
of `undefined` arrived as the word **"undefined"**, a JavaScript value name
offered as the whole explanation of why somebody's merge failed. An Office.js
async failure is routinely a plain object carrying exactly `name`, `message` and
`code` rather than an `Error`.

`formatValue` in `trace.ts` already refuses to print "[object Object]", with the
reasoning written beside it — "a line that occupies space and answers nothing".
This is the same rule on the path that reaches a PERSON rather than a log.

Every branch now answers something a reader can act on or repeat to somebody who
can: the message if there is one, the shape if there is not, and a sentence
saying the raise was empty if it was.

A test pinned `readable(undefined)` to the string "undefined", so this is a
deliberate change to an asserted expectation rather than an unnoticed gap. The
assertion is replaced and says why.

### Fixed — the slide-number boxes read things nobody typed

`Number` reads far more than anybody types into a slide box: `0x10` is sixteen,
`0b11` is three, `0o17` is fifteen, `1e2` is a hundred.

Both halves of that were wrong. **`0b11` was accepted**, so the pane quietly
offered to merge slides 3 to 9 for somebody who had typed neither number into
either box. And the refusals named a cause they had invented: `0x10` produced

> The block ends before it starts: slide 16 to 9.

about a slide 16 that appears nowhere on the user's screen and in nothing they
typed. The sentence that WOULD have been true — "Slide numbers are whole
numbers, and `0x10` is not one" — sat two lines above and never fired.

The text has to look like a decimal number before `Number` is asked what it is
worth. A fractional part of zeros stays admitted, because `4.0` is a whole
number and refusing it with "4.0 is not one" would be a false sentence.

Third time `Number()` has been wider than what a person typed in this codebase,
after `0x10` read out of a data cell as sixteen and a grouping the parser could
not read. The rule is the same each time: ask the shape first.

### Fixed — the number above the merge button ignored the conditions

It read "9 slides added after slide 10, leaving 19 slides in the deck." The plan
built eight.

The count was slides-per-record times rows, and a conditional slide is not
produced for every row — so the sentence a user reads to decide whether to press
was over by one for every slide a condition leaves out, and the deck size it
predicted was wrong with it. Two more places had the older half of the same
problem: the block subtitle and the row arithmetic still counted the rows as
PASTED after the user had unticked some, which the button itself stopped doing
some time ago.

`plannedSlides` counts with `slideApplies` — the rule `buildPlan` itself applies,
exported for the purpose — so the promise and the plan cannot answer
differently. The test asserts that agreement rather than the number, because the
number is only right for as long as the two rules are one.

`slidesAdded` is gone; it was the multiplication.

A note on the guard. The unit test on `plannedSlides` passed happily while the
card went on calling the old product — reverting the call site broke nothing
until a render test asserted the sentence itself. Both are there now.

### Fixed — a number too large for a double is now text on both sides

A correction to the entry above. Making `looksLikeNumber` ask `numericValue`
was described as behaviour-preserving, and that was true of the test suite and
not of one input class the suite does not cover.

The pattern matches any run of digits; `Number()` gives up above about 1.8e308.
So a 310-digit cell was typed `number` by the detector and refused by the
converter — the exact disagreement that change was made to end, sitting in a
shape nobody had written a case for. It is typed `text` now, and both halves say
so.

Text is the honest answer rather than the ideal one. A spreadsheet would call it
a number; this engine cannot hold it as one, and claiming it is a number and
then failing to format it is the worse of the two.

### Fixed — four spellings of "the block moved", and no two agreed

A pane update that sets one field is readable where it happens. One that clears
six is a RULE, and this rule was written out at four call sites:

| Moving the block by | cleared |
| --- | --- |
| typing in the slide-number boxes | block, fields, conditions, added, fieldNote |
| picking a selection | block, fields, conditions, added |
| a read that refused | block, fields |
| — any of them | never `imageFields` |

So picking a new selection left a note saying `{{Region}} put on the slide`
about slides the new block does not name, and a refused read left conditions
keyed to slide numbers the block no longer covers. `imageFields` was added in
the picture-picker fix and no path cleared it, so the picker could offer a
column the new block never mentions.

`blockMoved` and `dataChanged` in `src/pane/transitions.ts` are those two rules,
pure and stated once. `notice` is deliberately not in either: two callers set
one of their own, and clearing it inside would make the order of two statements
decide whether the user sees the sentence.

An architecture guard fails if `main.ts` clears the block or the row filter by
hand again — a fifth spelling written next to the fourth is how this started.

This is the part of the planned state refactor that is behaviour-changing on
purpose. The 28 `draw()` calls and the event loop are untouched: collapsing
those changes render TIMING, which the jsdom tests do not cover exhaustively,
and it is not worth the risk for a readability gain.

### Changed — a slide's relationship types are spelled out once

These strings decide which parts a clone copies, which a clone drops, and which
a removal deletes. They were written out across six files, three of them twice:
`NOTES_REL_TYPE` in `pkg.ts` and again in `clone.ts`, `COMMENT_REL_TYPES` in
both, `TAGS_REL_TYPE` in `clone.ts` and again in `tags.ts`.

The copies agreed and nothing had gone wrong. What made it worth ending is which
decisions they drive: one copy of the comment list says what a CLONE drops, the
other says what a REMOVAL deletes. PowerPoint has already added a second
spelling of comments once — the modern web one, under a Microsoft namespace —
and adding a third to one copy and not the other leaves a clone carrying a
comment part the removal will not clean up, or a removal deleting one a
surviving slide still points at.

`parts.ts` holds the vocabulary, and the two ownership patterns with it: what a
slide owns, and what a chart or diagram may drag out of the package. Both were
moved rather than retyped, so the anchoring that keeps a crafted relationship
from naming `ppt/presentation.xml` is character-for-character what it was.

An architecture guard fails on any relationship type written out anywhere else.

Behaviour-preserving: 859 tests before, 859 after, none edited.

### Changed — one list of the parts a merge touches

`prepare` reads those parts to say what a block holds; `runPlan` and
`mergeGraphics` write into them. Each side assembled its own list, from its own
imports, in its own order — and `prepare.ts` carried the rule they were supposed
to obey as a COMMENT: "the one rule that keeps it from happening a third time is
that this list and `runPlan`'s are the same list".

It happened three times. Speaker notes were merged and not scanned, so a block
whose placeholders lived there was refused as empty. Then chart labels, the same
way. Then a chart's value cells, which live in a workbook the scan never opened.
Each was fixed by adding the missing part to one of the two lists, which fixes
the instance and leaves the class.

`fieldSites` is the list now, and both sides read it. A part type missing from it
is invisible to both rather than to one — which turns a silent wrong answer into
a visibly missing feature.

Two lookups collapsed into it. `chartWorkbooksOf` and `workbookOfChart` were the
same walk of a chart's relationships, one taking every package and one taking
the first, and `mergeGraphics` ran both — so every chart's relationships were
walked twice per record. A site carries the list, the numbers pass takes the
first for its pairing and the text pass takes the deduplicated set, and both
functions are gone.

Two architecture guards hold it: the three merging files may not import
`graphicPartsOf` or `notesPathFor`, and no file outside `sites.ts` may call
them.

Behaviour-preserving: 857 tests before, 857 after, none edited.

### Changed — a value gate is now defined as the parser behind it

`detectType` asks whether a cell is a number or a date; `applyFormat` asks for
its value. Those two answers came from patterns in different files, and they
drifted three times in one day — `Number()` accepting `0x10` that the pattern
refused, a widened date gate against a private copy of the pattern inside
`parseDate`, and a grouping pattern admitting `1,234,5` that `numericValue`
could not read. The symptom was the same every time: a column typed one way,
converted another, half formatted and half raw with nothing saying why.

The shapes now live in the file that parses them. `looksLikeNumber` IS
`numericValue(v) !== undefined`, so for numbers the disagreement is no longer
possible rather than merely swept for — the 6190-case sweep still runs and is
now a tautology, which is the right end state for a property that should be
structural.

Dates keep their asymmetry, deliberately and visibly: a date can be well formed
and impossible, so `31 Feb 2026` must pass the gate and fail the parse. That
shape test is `dateShape`, exported from `format.ts` beside the parser that is
its only other reader.

Two architecture guards hold the direction: `format.ts` may not import from
`recordset.ts`, and `recordset.ts` may hold no value pattern of its own.

Behaviour-preserving: 855 tests before, 855 after, none edited.

### Added — a package-validity case carrying every feature at once

The cases there take one feature at a time, which is right for saying WHICH one
broke. It left the combination untested, and the combination is where part
numbering, relationship ids and content types can collide with each other rather
than with themselves.

Pictures were the piece none of them had. A media part is declared by EXTENSION
rather than by name, so it is the one part type that can conflict with a
declaration the deck already carries, and deduplication means the number of
media parts is not the number of rows. The case also carries a row whose picture
nobody supplied.

**No defect found.** The package is legal after the merge and after the sweep.

### Fixed — an author who wrote `{{Photo|image}}` got a pane with nowhere to attach the files

The pane decided what a picture was from the DATA's detected types. The engine
decides from the FIELD's format. They disagreed about every column the type
detector had turned down.

`detectType` is all-or-nothing on purpose — one cell reading `n/a` in a column of
file names makes the whole column text, so a column of `.svg` names is not
offered as pictures and then failed one row at a time. That part is right and is
unchanged.

What was wrong is what followed. The picker appears only when the data refers to
pictures, and that list came from the detected types alone, so a column one
stray cell had kept out of the type meant:

- no picker, and therefore no way to attach the files at all;
- an insert button writing `{{Photo}}` rather than `{{Photo|image}}`, which
  merges the file name onto the slide as text;
- a merge that left every picture placeholder standing, silently.

The engine placed the picture perfectly well in that state — `placed: 1` — when
files were supplied. Nothing in the pane could supply them.

`docs/MANUAL.md` documents `{{Photo|image}}` as the way to ask for a picture, so
an author writing it by hand is the documented path, not a workaround.

`prepare` names the picture fields now — it reads the slides already — and they
travel to the pane beside the ordinary ones. A column is a picture column when
the detector says so OR when a field on the slides asks for one.

`imageFieldsIn` has answered "which fields ask for a picture" since it was
written. Until now nothing in the product called it.

A consequence worth stating: in an author-declared picture column, a cell that
is not a file name — the `n/a` that caused all this — is now listed among the
pictures the data asks for and reported as missing. That is what the ENGINE does
with it too, and the two agreeing is the property that matters.

### Fixed — a deck that grew by more than the package held authorised deleting all of it

`added` is measured from the DECK rather than from the plan, and that is right:
when the host lands fewer slides than it was handed, the deck knows and the plan
does not. It was wrong in the other direction, and not cosmetically.

`sweepPlan` refuses to sweep when the deck grew by more than the run added —
that clamp is what keeps an undo off a stranger's slides. An uncapped `added`
absorbs the excess, so `grew` and `added` are equal by construction and the
clamp can never fire. Six slides arriving across an insert of three would have
authorised deleting six, three of them somebody else's.

`added` is capped at what the package actually held now, so the same case leaves
`grew > added` true at undo time and the sweep refuses — which is the answer
that rule was written to give. A deck that shrank reports zero rather than a
negative count.

### Fixed — `insertVerdict` called two impossible things a partial insert

The branch for "the delta is neither zero nor what was expected" also caught the
delta being LARGER, and the delta being negative:

- `5 of 3 slide(s) landed, which is a partial insert rather than a refusal`
- `-2 of 3 slide(s) landed, which is a partial insert rather than a refusal`

Neither is a partial insert, and the first is a sentence that cannot be true. It
named a cause for a condition it had not distinguished.

Both are `unknown` now — slides plainly arrived — with a detail that names the
condition and stops: the deck grew by more than the package held, so this run
cannot say which of them are its own.

Whether a host can land more than it was given is unproven. A user clicking in
PowerPoint while the pane works, or another add-in, reaches the same state, and
the clamp it defeated is the one guarding somebody's slides.

### Fixed — a removed slide left its tag part behind

`orphanedParts` collects what only the departing slide keeps alive, and it
collected charts and diagrams. Tag parts were not on the list: `writeSlideTags`
writes `ppt/tags/tagN.xml` per slide, exactly one slide points at it, and it is
unreachable the moment that slide goes. Every removed slide left one behind —
with a content-type override and nothing referring to it, which is the shape
this file already chases for notes pages and for comments.

It reaches further than a swept preview. On the `file` route — every host below
PowerPointApi 1.10 — the package IS the user's whole presentation, and every
slide that is not a clone is removed from it before the insert. A deck whose
slides carry tags, this add-in's own from an earlier merge or another add-in's,
shipped one orphan per slide back into their deck.

Found by sweeping rather than reading: merge three records from a slide carrying
a chart, a workbook, SmartArt, notes and tags, remove every merged slide, and
require that nothing under `ppt/charts`, `ppt/diagrams`, `ppt/embeddings`,
`ppt/notesSlides` or `ppt/tags` is left unreachable. 57 parts went to 27 with
three tag parts standing; they go to 24 now. The sweep is a test, and the part
type that was missed is by definition the one nobody writes an assertion for.

The name is anchored like the others, so a crafted relationship cannot point the
sweep at a part the presentation needs, and a tag part another tool named
something else is left alone.

### Fixed — a selection naming one slide twice could merge a slide nobody picked

`blockFromSelection` decided contiguity by comparing `to - from + 1` against HOW
MANY numbers the selection produced. A count is not an alignment: name one slide
twice and the count covers a gap it was there to catch.

Select slides 1 and 3, with slide 1 named twice: three numbers spanning three
slides, so it came back **"slides 1 to 3"** and put slide 2 into the template
block. That function's own comment says a gap is refused "because closing it up
would silently add slides the user did not pick" — and this closed one up.

The other direction was merely annoying: slides 1 and 2 with one named twice
was refused as not contiguous, and so was a single slide named twice.

Slide numbers are collected once each now, and the count means what the check
reads it as.

Whether a host ever names a slide twice is unknown — `getSelectedSlides` is not
documented to. It is the same API whose ids are not roundtrippable
(office-js#2474, repaired a few lines above), which is reason enough not to
leave a count standing in for the check.

### Fixed — a slide whose only field was its chart's number was refused as empty

`prepare.ts` states the rule: "this list and `runPlan`'s are the same list." A
scan that reads fewer parts than the merge writes refuses a block it would have
merged, and tells the author to go and type field names onto a slide that
already carries one.

That has now happened three times. Speaker notes, then chart labels, and now
chart VALUES — the first two are written up in that same file, a few lines above
where the third one was.

A value cell holds its placeholder in the **workbook**. `fieldsIn` cannot see it,
because `<c:numCache>` is deliberately left out of the text pass — a formatted
number is unplottable — and the workbook was not read at all, on a rationale
that was true right up until chart numbers became a feature: "its strings are
the same strings as the chart's own cache, so reading it would name nothing
new."

So a slide carrying a chart whose numbers vary per row, and nothing else, was
told it "has no {{fields}}, so every copy would be identical". The merge fills
it: two numbers across two slides, in the case that found this.

The scan is a dry run of `mergeChartNumbers` itself now, driven by a resolver
that records each name and answers null — the same answer a placeholder with no
column gets everywhere else, which leaves the text standing, writes nothing, and
repacks no workbook. The same WALK rather than a second reader, so the scan and
the merge cannot hold different opinions about which cells carry a placeholder.

That "writes nothing" is a claim about a code path, so it is measured at the
only two places that path could write: the embedded workbook's bytes and the
chart's cached values. Both are asserted unchanged.

The labels still come from the chart's own cache rather than the workbook. Those
two hold the same strings, and reading both would name nothing twice.

### Changed — `sweepPlan` checks the count it deletes by, and states its rules as properties

`sweepPlan` decides which slides an undo removes from somebody's presentation.
It integer-checked the two deck counts and trusted the third — `added`, the one
that decides how many slides come out.

`added: NaN` walked every clamp untouched. `grew > NaN` is false,
`Math.min(NaN, grew)` is NaN, `NaN <= 0` is false, `NaN < deckAtStart` is false,
and `{ from: NaN, count: NaN }` came back out of the function whose whole job is
refusing to produce a plan it cannot justify.

Nothing can reach it today: `added` is a step count and the types carry it. It
is checked because the other two are, and because a reader comparing the three
has to see one rule rather than work out which one is trusted.

The rules are now also asserted over the whole space — every combination of the
three counts up to 12, 20 and 12 — rather than at the points somebody thought
of: a plan may not reach a slide the user owned, may not reach past the end of
the deck, may not remove more than the run added, may not be empty or
fractional, and may not exist at all when the deck grew by more than the run
added. **No violation was found.**

That last property is there because the first four cannot see it. A plan built
after a stranger appended slides still starts past `deckAtStart` and still ends
at the deck's end — it satisfies every structural rule and is still somebody
else's slides. Removing the guard for it left the four happy; only the stated
property and an existing point test noticed.

### Fixed — step 4 named the wrong thing to do while a preview was on the slides

`blockedReason` promises that every sentence it returns "names the thing the
user has to do". Preview a row, go back to step 1 and type in a slide-number
box: that clears the committed block — deliberately, it is stale — and does not
end the preview. Step 4 then read **"Choose the slides that repeat first."**
directly above a working **"Remove the preview"** button.

Never stuck. Navigation between steps is ungated on purpose and the button
always worked; it was the sentence that was wrong, at the one moment the user
had an obvious right thing to do. A live preview is not a blocked step, because
there is something to do on it and the button already does it.

Found by sweeping 576 combinations of the things a step gates on — block, rows,
fields, previewing, running, added, excluded — against the rule that a blocked
step must never offer a pressable button. 64 states broke it, all this one. The
sweep is now a test.

What it does NOT catch, said plainly: a guard that goes missing makes the
sentence and the button AGREE, so the invariant stays true. It sees a wrong
sentence over a live button, not a missing sentence.

### Added — a sweep for per-record parts two slides both point at

Every part a merge writes into has to be the merging slide's own. A part type
with no branch in `cloneSlideGraphics` means every record merges into ONE part
and the whole deck shows the last row. That has been found three times, in three
different part types, and each time the fix was a named branch for the type that
was missed.

The new test asserts the property instead: merge three records from a slide
carrying a chart, its workbook, SmartArt, speaker notes and tags, walk every
part reachable from each output slide however many hops out, and require that no
two of them share anything under `ppt/charts`, `ppt/embeddings`,
`ppt/notesSlides`, `ppt/tags` or a diagram's `data`/`drawing`. Reachability
rather than a list of branches, so a part reached by a hop nobody thought about
is still covered. Both places a diagram's drawing can be hung are swept.

Deliberately not covered: layouts, masters, themes, and a SmartArt's `layout`,
`colors` and `quickStyle`. Those are static definitions no pass writes into, and
sharing them is as right as sharing the theme. Media is shared on purpose too —
one logo across 240 rows is one part.

**No defect found.** The sweep passes as written, and it was measured rather
than trusted: removing the chart branch, the diagram branch, or the workbook
clone each makes it fail, the last being the subtle one where the chart is
copied and the workbook behind it is not.

### Fixed — a foreign tag's whitespace did not survive a merge

An XML parser NORMALISES an attribute value: a literal newline, carriage return
or tab inside one is read back as a **space**. `xmlAttr` escaped the five markup
characters and wrote those three literally, so a tag value carrying any of them
came back on one line after a single merge — and stayed at that wrong value
afterwards, which is the shape nobody reports.

It cannot reach our own tags; a run id and a record number have no whitespace in
them. It reaches a FOREIGN tag, which `mergeTagPart` carries through untouched
and `docs/MANUAL.md` promises survives a merge. Another add-in keeping anything
formatted in a tag got it back flattened.

This is the second half of a defect already fixed once from the entity side: the
values were being escaped exactly once, and three of the characters that need
escaping were not on the list.

### Fixed — a chart's numbers were found by file name rather than by declaration

The embedded workbook's worksheets were collected by matching part names against
`xl/worksheets/sheetN.xml`. Excel writes that name. The format does not require
it, and a workbook built by anything else need not.

The result was the quietest failure in this engine: not a wrong number and not a
refusal, but **nothing**. `filled: 0`, `refused: 0`, the chart keeping its cached
values, the deck looking finished — and the placeholder still sitting in the cell
for whoever eventually clicks Edit Data.

A workbook states which parts are its sheets, in `xl/workbook.xml` and its
relationships. That is read now. Two things came with it:

- The worksheet list is resolved **once per workbook**. It was resolved once per
  series, so both of those parts were reparsed for every `<c:numRef>` of every
  chart of every record.
- The relationship target goes through the package's own `resolveTarget`
  instead of having `xl/` glued on by hand — right for the target Excel writes,
  wrong for the two other shapes a target is allowed to take.

### Changed — `fieldPattern()` replaces the exported `FIELD` regex

`FIELD` is global, so it carries `lastIndex` between calls: `matchAll` copies
that index onto the matcher it builds, and `test` leaves it wherever the match
ended. One caller's leftover state decides where the next one starts reading,
and the second silently misses every field before that offset.

Three call sites had already worked around this by hand with
`new RegExp(FIELD.source, FIELD.flags)`, a fourth reset `lastIndex` around each
use, and a test wrote the clone a fifth time. Five spellings of one precaution
is how the one that forgets gets written.

`FIELD` is no longer exported. `fieldPattern()` hands back a fresh matcher, and
the compiler now enforces what those five hand-written precautions were for.

### Fixed — a condition now decides before an empty cell does

`onEmpty: "skip"` drops a record whose fields are not all filled. It looked at
the fields of **every slide in the block**, including the ones that record's own
conditions had already left out — so a customer with no renewal note vanished
from the deck entirely, over a blank cell on a renewal slide they were never
going to get.

Conditions are evaluated first now, and the empty-field check sees only the
slides the record will actually receive. The policy itself is unchanged: a blank
cell on a slide the record IS getting still drops it.

A record dropped whole no longer also reports the slides its conditions left
out. It contributed nothing; saying it contributed two absences as well would be
two answers about one record.

Latent rather than shipped. The policy reaches `buildPlan` through the office
merge request and the pane does not set it, so nobody can have hit this — it was
a trap laid for whoever wires it up, and the failure it produces is a record
silently absent from the deck.

### Fixed — the number gate admitted a form the number parser cannot read

`detectType` asks `looksLikeNumber`; `applyFormat` asks `numericValue`. The
grouping pattern behind the first accepted `1,234,5` — a separator used for
grouping AND for the decimal — and the second returned nothing for it. A column
holding one typed as a number and rendered raw, which is the disagreement this
pair of functions exists to prevent and the second time it has happened.

The pattern captures the grouping separator now, makes the later groups repeat
THAT one, and forbids a decimal part from reusing it. A sweep over 6190
arrangements of digits and separators asserts the two agree about every one;
under the previous pattern 64 of them disagree.

**`1.234` reads as 1234, and the comment beside it said otherwise.** It claimed
"a single `1.234` stays a decimal". It never did — the quantifier is one or
more, so a lone three-digit group has always been read as grouping, which is
what the comma branch does with `1,500` too. The reading is right and the
symmetry is right; the comment was a leftover from an earlier intent, and the
two readings differ by a factor of a thousand. It is written down and pinned by
a test now.

### Fixed — two header shapes the image reader got wrong

**A JPEG padded with fill bytes was reported unreadable.** Any number of `0xFF`
bytes may sit between segments, and the standard says so. Read as a marker
itself, `FF FF C0 …` takes the frame header's own first bytes as a segment
length, skips a nonsense distance and runs off the end. A perfectly good photo
came back as nothing, and the pane reports that as an unreadable file — one
placeholder left visible with nothing to say why.

**A BMP with the 12-byte OS/2 header was measured wrong, and placed anyway.**
That header keeps width and height as two 16-bit numbers at offsets 18 and 20;
every later header keeps them as two 32-bit numbers at 18 and 22. Reading the
second shape out of the first does not fail — `200 | (100 << 16)` is a number,
so a 200 × 100 bitmap measured 6553800 × 1572865 and was cropped to a ratio with
nothing to do with it. Nothing anywhere said the size had been invented.

The reader now asks the header how long it is, which is the field at offset 14
it was skipping. A header length it does not know is refused rather than read at
the wrong offsets.

The existing top-down-BMP test built its fixture with that field left at zero,
which is not a BMP any encoder writes; it passed only while the reader ignored
the field. The fixture states a length now. Its own claim — that a negative
height is a direction and not a size — is untouched.

EXIF and progressive JPEGs, the `C0`–`CF` markers that are not frame headers,
and a truncated frame were all checked at the same time and were already right.

### Fixed — the picture pass took the rest of the paragraph with it

Placing a picture blanked **every text node in the paragraph**, not the
placeholder's own characters. Whatever shared the paragraph went with it.

A caption disappearing is at least visible — somebody notices the slide is
bare. The case worth the fix is `{{Name}} {{Photo|image}}` in one line: the
picture lands, the merged NAME is wiped, and the slide looks finished on every
copy with nothing in any count saying otherwise.

Three more defects fell out of the same loop.

**A shape has one fill.** A second picture field in the same shape overwrote the
first, counted itself into `placed`, and left a media part and a relationship
behind for a picture that is not on the slide — a count saying two where the
deck shows one. The first wins now, the rest are reported as `crowded`, and
their placeholders are left standing so the author can see which was not drawn.

**A table cell was skipped in silence.** The file said otherwise: "A field whose
text is not inside a `<p:sp>` at all — in a table cell, say — is reported as
missing rather than guessed at." The walk started at `<p:sp>`, so a table's
paragraph was never visited and the check for it sat below a loop that could not
reach it. Dead code under a sentence promising the behaviour. The walk goes over
paragraphs now and asks each which shape it is in, which makes the documented
answer the real one.

**`docs/MANUAL.md` was wrong about URLs.** It said a cell holding a URL "is not a
picture and is merged as text". A column of `https://…/ada.png` is typed as an
image column, and the name at the END of the URL is matched against the files
you picked — so it finds `ada.png` if you picked it, and nothing if you did not.
The manual now says that.

The span arithmetic the text pass already used is now shared rather than
approximated: `editRuns` replaces characters by offset into a paragraph's joined
text, and both passes call it.

### Fixed — one definition of a number, and one of a written date

Two pairs of functions each answered the same question two ways, and in both
cases the disagreement reached a slide.

**A number.** `detectType` asked a regex; `numericValue` finished with a bare
`Number()`, which is a far wider gate. `Number()` reads `0x10` as sixteen,
`0b11` as three, `0o17` as fifteen and `1e3` as a thousand. A column of product
codes was therefore typed as text — so the pane never offered it as a number —
and then converted anyway wherever a format spec reached it. A code that turns
into `16` across a merged deck reads as deliberate, which is what makes it worse
than a cell left alone. `numericValue` now asks the same exported gate
`detectType` does, and refuses what a spreadsheet would refuse.

The comment inside `numericValue` had already recorded this shape once, in a
merge where "half of it rendered formatted and half rendered raw". It was the
same bug, still open, in the function that described it.

**A written date.** `NAMED_DATE` allowed exactly one separator character between
the day and the month name. Danish writes the day as an ordinal, so
`1. marts 2026` — the ordinary long form — is a period AND a space, and was
refused, while `1 marts 2026` and `1.marts 2026` were admitted. The month-name
table added for Danish dates was reachable mainly by spellings nobody types.

Widening that gate alone made it briefly worse: `parseDate` carried a private
copy of the same pattern, so the value typed as `date` and then rendered raw —
the two-renderings failure again, entered from the other side. Both now match
with one exported regex, and a test asserts the pair agrees for every value that
is both well formed and real. Nothing was loosened about ambiguity: `03/01/2026`
is still refused, because a month spelled out is unambiguous however it is
punctuated and a slash date is not.

### Changed — a placeholder with no column warns instead of refusing

The merge step used to refuse to run at all while any placeholder on the slides
had no column behind it. It names them and runs now, and the sentence says what
will happen — `No column for Nickname. It will stay on the slides as written` —
rather than what to fix.

Three parts of this project already disagreed with the refusal. The **engine**
leaves such a placeholder on the slide, deliberately, so a half-filled deck does
not look finished. The **preview step** had no such check and ran the ordinary
merge with one, correctly. And **docs/MANUAL.md** promised it in as many words:
"a row whose picture is missing keeps its placeholder, exactly as a text field
with no column does".

The asymmetry was the tell. A field whose COLUMN was missing was refused; a
field whose column existed but whose PICTURE was missing was allowed, and
documented. Both end with a placeholder on the slide.

What the refusal protected against is real — a typo merged across 240 slides is
expensive — but by that screen the user has been told twice: the fields step
outlines the chip and names it in a card, and this sentence sits directly above
the button. Being told is the protection; being stopped was not.

It is also what made `docs/TEST-KIT.md` step 5 possible to follow. The kit's own
template carries `{{Nickname}}` on purpose, so its documented run could not be
completed in a real host — the round of 2026-08-28 reached a merge only by adding
a column the kit deliberately omits.

### Added — a chart's numbers, per recipient

Type `{{Revenue}}` into a value cell of the chart's own data sheet, through Edit
Data. Each merged copy then gets that row's figure, so the bars differ per
recipient.

No syntax of its own, which was the open question. The backlog assumed one would
be needed — "there is nowhere in the values to put one" — because a
`<c:numCache>` cell has to parse as a number. There is somewhere: the embedded
workbook's own cell, where the placeholder is an ordinary shared string and
`{{Column}}` means what it means everywhere else. `<c:f>` joins the two:
`Sheet1!$B$2:$B$3` beside a two-point cache says point 0 is B2.

Both copies move together, which is the whole difficulty. The workbook cell goes
back to being numeric, because a chart plots nothing from text and Excel shows
whoever presses Edit Data what is really in there; the chart's cache is written
too, because that is what PowerPoint draws from without opening the workbook.
Filling one and not the other is the half-merge this project already knows from
the label side — right until Excel touches it, then reverted in front of the
user.

Two deliberate refusals. A format is ignored in a value cell:
`{{Revenue|number:0}}` would hand back `1 250 000`, which is the right string and
an unplottable cell. And a placeholder that does not resolve to a number is left
exactly as written rather than becoming a zero, because a zero is a bar the data
never asked for. The run counts those.

### Added — a test kit for the real-host round

`test-kit/` holds a template carrying a chart, a picture frame, formatted
numbers and dates, speaker notes and a field with no column, plus three rows and
three photos; `docs/TEST-KIT.md` is the checklist, and `test-kit/PROMPT.md` is
the prompt for driving it from a local Claude Code session.

The template's chart was authored by python-pptx rather than by this project.
That is the point of it: every other fixture here was written by the same author
as the reader, and a reader built from the same misreading of a chart part
agrees with itself perfectly. The SmartArt is left for the tester to add in
PowerPoint for the same reason — nothing outside PowerPoint can author one, and
a diagram PowerPoint wrote is the stronger test.

`test/test-kit.test.ts` merges that committed template on every CI run — a
foreign-authored chart in the suite, and a guarantee that the deck a person is
asked to open has not quietly stopped merging between rounds. What it cannot do
is open the file in PowerPoint, which is what the round is for.

### Added — charts and SmartArt are merged

A placeholder in a chart title, in its category labels, or in a SmartArt box is
filled like any other. Each merged slide gets a chart of its own, so the copies
differ — which is the point, and is why a merged deck with charts weighs more
than one without.

Their text was never unreadable: it is DrawingML, the same `<a:p>` and `<a:t>` a
slide holds, and `fieldsIn` has reported these placeholders for as long as it
has existed. What stood in the way was four facts about where the same string is
kept.

**The parts are shared by every clone.** `cloneSlide` copies a slide's
relationships wholesale, so all 240 copies point at the template's own
`chart1.xml`. Merging into that writes one record's values and shows them on the
whole deck. Notes pages had this defect and were fixed by cloning; comments had
it and are dropped. The rule the three make: a part a merge writes into may
never be shared by two slides. What stays shared is the read-only styling — a
chart's colours and style, a diagram's layout, quick style and colours — because
copying those per record multiplies a template's styling by the row count for no
change in what anybody sees.

**A chart's labels are not paragraphs.** The category and series names a user
actually writes live in `<c:v>` inside a `<c:strCache>`. Not every `<c:v>`: the
same element holds the chart's NUMBERS inside `<c:numCache>`, where the content
has to parse as a number, so a merge that took them all would replace a bar's
height with "Nordics" and produce a chart PowerPoint reads as corrupt.

**A chart's labels are also in the workbook behind it.** The cache is what
PowerPoint draws; the workbook is what Excel opens on Edit Data, and closing
that Excel refreshes the cache from the workbook. Merging the cache alone gives
a deck that is right until somebody clicks the button and watches the labels
revert to `{{Region}}`. So each copy gets its own workbook and its shared
strings are merged too — a package inside the package, opened with its own zip
reader. One that cannot be opened is reported rather than thrown on: the chart
still merged, and losing 240 slides over an embedded object another tool wrote
is the wrong trade.

**SmartArt keeps its text twice.** `dataN.xml` is the model and `drawingN.xml`
is the laid-out rendering PowerPoint puts on the screen. Merging the model alone
produces a deck whose SmartArt still reads `{{Name}}` to every viewer. The
drawing hangs off the data part rather than off the slide, which is why cloning
what the slide names is not enough.

One reader now finds all of it — a DrawingML paragraph, a chart's cached string,
a workbook's shared string — so what the pane counts and what the merge fills
cannot come apart.

### Changed — a chart's placeholder is an ordinary field

It was reported as unfillable, in the fields list and in a refusal when a block
had no other placeholder. Both are gone: the merge fills them, so they are
counted as fields, and a block whose only placeholder is in a chart merges. A
chart fill counts into the "N placeholders filled" line for the same reason —
its zero is the alarm that says a merge added every slide and filled nothing,
and a template whose fields all live in a chart would otherwise raise it about a
merge that worked.

### Fixed — a removed template slide no longer strands its chart

The template slides are taken out of the package before it is handed to
PowerPoint. Their charts used to be shared with the copies and so stayed
referenced; now each copy has its own, and the template's would have been left
in the file with nothing pointing at it — a chart and an embedded workbook per
template slide. They are swept, and only when no other part references them,
which is what keeps a diagram's shared layout from being swept with them.

### Added — picture fields

A cell can name a picture file, and the shape the field sits in is filled with
it. `{{Photo|image}}` in a rectangle on the template puts the row's picture in
that rectangle, at the template's size and position, on every merged slide.

Three fits, because there is no one right answer: `image` crops to fill (the
default, and what a page of portraits wants), `image-fit` letterboxes the whole
picture inside the shape, `image-stretch` squashes it to the shape's exact
proportions.

The files come off the user's own disk through the browser's file picker, on
step 2, which grows the picker as soon as a column's cells look like file names.
Nothing is uploaded and nothing is fetched: a sandboxed cross-origin task pane
has no other route to a local file, and for a merge whose premise is that the
data does not leave, no other route is wanted. A cell holding a URL is text.

Matching is by base name, ignoring folders and case, so `Photos\\ada.PNG` in a
spreadsheet finds `ada.png` from the picker — the same rule in the pane's tally
and in the engine, computed by one function, because a pane counting matches by
a different rule than the merge uses would promise pictures that never arrive.

A row whose picture is missing keeps its placeholder, exactly as a text field
with no column does, and the pane names what it has not got rather than counting
it. A file whose bytes are not the picture its name claims is left out and
named, not written into the deck.

Written into the package as an ordinary `<a:blipFill>` on the shape: one
`ppt/media/imageN.ext` per distinct file however many rows use it, one
`Default` content type per extension, and a relationship per slide that
references it. The crop is `<a:srcRect>` and the letterbox is
`<a:stretch><a:fillRect>` — both insets in thousandths of a percent, both
derived from the shape's box and the picture's own pixel dimensions, which are
read from the file's header rather than trusted from its name.

### Fixed — a comment on the template landed on every merged slide

Found by answering probe question 5 on a deck that finally carried comments.
`exportAsBase64Presentation` drops comments and `ppt/authors.xml` outright —
four parts in, none out, so office-js#6867 reaches the presentation-level call.

That is harmless in itself, and it exposed something that was not: **the two
template routes disagreed.** The subset route (1.10) produced comment-free
clones because the host had already dropped them; the file route — every host
below that — copies the slide's relationships wholesale, so every clone got a
relationship to the TEMPLATE's comment part. Three slides, one
`modernComment_101_AEAB9DA1.xml`, measured before anything was changed. A
reviewer's "check this with Legal" would appear on all 240 merged slides, as one
shared thread.

Copying the part per clone would be worse rather than better: the same note 240
times, deliberately. A comment is an annotation about the template, not content
the template produces — which is the answer the 1.10 host had already chosen, so
dropping them is also what makes the two routes agree.

`cloneSlide` drops comment relationships in both spellings — the classic
`commentN.xml` and the modern `modernComment_<id>_<hash>.xml` PowerPoint on the
web writes under a Microsoft namespace — and `removeSlide` now takes a comment
part away with its slide, as it already did for notes, so removing the template
on the way out cannot strand one.

The user's own comments are untouched: this drops a COPY's inherited reference,
never the template's own.

### Fixed — a clone could share the template's notes page, and ship the wrong record's notes

Found by a bug hunt and reproduced on real bytes before anything was changed.

Part names in a package are arbitrary, and the slide and notes-slide sequences
drift apart the moment a slide is deleted — so a one-slide deck can perfectly
well keep its notes in `notesSlide2.xml`. `cloneNotesSlide` named the copy after
the SLIDE number, which lands straight on that part. Nothing complains:
`copyPart` overwrites silently and `addContentTypeOverride` no-ops on an
override that is already there, so the package stays structurally valid and is
wrong in two ways at once.

- **The second record's slide shipped the first record's notes.** The clone
  shares the template's notes page, that page is merged, and the next clone then
  copies notes whose placeholders are already gone.
- **The package went out with a notes relationship pointing at nothing.** The
  template is removed on the way out — that is how the clones end up alone in
  the package — and removing a slide takes its notes page with it. Shared, that
  page was the clone's too. PowerPoint reports this as a damaged file without
  saying which part it could not find.

Notes parts are numbered from their own sequence now. `nextTagNumber` had the
right shape all along, one file over: ask whether the path is free rather than
assume it.

### Fixed — the crash record was write-only in the window it exists for

`merge()` writes the crumb BEFORE handing the package to PowerPoint, with
`added: 0`, because a tab that dies during that call never comes back to write
the real number — and that is the whole reason the file exists. `readCrumb`
refused a zero, so the record was readable only after the run it was insurance
against had already succeeded.

Zero authorises nothing and must not: `sweepPlan` refuses a count of zero, and
deriving one from the deck's growth would sweep whatever has been appended
since. So the pane TELLS the user — a merge did not finish, the deck had N
slides before it and has M now, check the end of the deck — and never offers a
delete it cannot clamp. Said once, then the crumb is cleared, because there is
no action attached to it.

### Added — the JSON manifests are validated, and a release checks its own URLs

Two gaps between "green repo" and "a stranger can install this".

**`manifest-prod.json` is a release asset that no external authority read.** CI
and the release both ran Microsoft's validator over the XML manifests only,
while the release note points administrators deploying to a whole tenant at the
JSON. Measured rather than assumed before wiring it up: the validator reads a
JSON manifest against its schema offline, needs no service for it, and exits 1
on one missing `id`, missing `version` and carrying a wrong `manifestVersion`.

**Nothing ever asked whether the host a manifest names is serving.**
`checkManifest` asks whether a production manifest points at localhost, which is
a different question and passes cleanly for a manifest pointing at a domain that
404s — an add-in that installs perfectly and shows a blank ribbon button and an
empty pane. The release job now fetches every URL the manifest names on its own
origin. On the release job only: a third-party outage must not block a merge,
which is the reasoning that already keeps the validator out of `test`.

### Fixed — a field name could not contain a space, so Excel's own headers were invisible

Reported from a real run, an hour after the Insert buttons shipped, on a deck
whose slides plainly carried `{{Min. of cost}}`, `{{Row Labels}}` and
`{{Sum of quantity monthly}}` — put there by those very buttons — while the pane
answered *"Slides 2 to 3 carry no fields yet."*

The reader's character class was a list of allowed characters and a space was
not on it. Those three are the literal default headers of an Excel pivot table,
which is the commonest thing anybody pastes into this add-in, so the engine was
blind to most real data. The rule is stated the other way round now, which is
what it always meant: a name is anything that is not a brace, a pipe or a line
break, and it must carry at least one letter or digit. `{{ }}` and `{{!!}}` are
still not fields; `{{a b}}`, which had been swept in with them, is one.

**The deeper defect was that the pane and the engine disagreed at all.** The
Insert button built `{{Column}}` and `FIELD` read it back, and nothing checked
that those two agree — so the button could put a token on the slide that the
reader could not see. `canBeField` is the engine's own reader, asked rather than
restated, and the Fields step now offers a button only for a column it can
actually write as a field. It names the ones it will not offer and why, because
the fix is to rename the column and an absent chip says neither which nor why.
The case that made it worth a shared function is not a header that fails to
match but one that matches a *different, shorter* name: `Total|EUR` would have
put a field called `Total` on the slide, bound to nothing, silently.

Also fixed on the same screen: a freshly inserted field put *"{{City}} put on
the slide"* directly above *"these slides carry no fields yet"*. An insert lands
on the slide and tells the pane nothing, so the second sentence was read off a
template read taken before it — the screen contradicting itself about something
the user had just done. The stale line is withheld while the note is up, and the
note already ends by asking for the read that settles it.

### Added — a button per column, and the steps in the order the work happens

Asked in these words after a first real run: *"how do I insert the fields? it
should be: select slides for repeat → paste data → insert fields → merge."*

The pane has five steps now — **Template, Data, Fields, Preview, Merge** — and
the third one hands you a button per column. Press it with the cursor in a text
box on a slide and `{{Column}}` is typed in there
(`Office.context.document.setSelectedDataAsync`, a Common API with no
requirement set, so nothing is declared and the call is guarded at run time).
Where the host refuses — usually because nothing on the slide is selected — the
token goes on the clipboard instead and the pane says so; where even that is
refused, the sentence carries the token to read off the screen. All three
outcomes name the token.

The reorder is what makes it possible. A field is a column name, so there is
nothing to insert until the data is attached: the old order asked for the fields
first and refused to go forward without them, which meant telling a first-time
user to go and type names they had no way to know. The template read tolerates a
block with nothing on it now (`allowEmpty`, passed by `inspectBlock` alone) —
`runMerge` still refuses one, because N identical copies is never what anybody
meant, and the pane refuses it too, before a host call is spent.

Nothing tells the pane that a user typed on a slide — there is no event for it —
so the fields step's own button reads *Check the slides for fields* and reads
them back.

The conditional-slide control moved to the merge step, beside the row picker:
"which rows" and "which slides" are the same kind of question, and the fields
step is now about putting placeholders onto slides.

The *Attach data first to see your column names* link on step 1 is gone with the
problem it worked around — a link that sent the user backwards through a wizard.

### Fixed — the word PowerPoint had already taken

The same run answered "Slides 2 to 3 has no placeholders" on a deck whose slides
held two of PowerPoint's own empty content boxes — which PowerPoint itself calls
placeholders. First contact, on an empty deck, reading as the add-in being
broken rather than as an answer.

The refusal names the syntax now and says where the names come from. A first
draft said "type something like `{{Name}}`" and was rejected by the same
reporter: at step 1 nothing is attached and nobody knows their column names yet.

### Fixed — a placeholder only in the speaker notes was refused

`runPlan` merges the notes page and always has. `prepareBlock`'s field scan only
read the slide, so a block whose placeholders lived in the notes was refused
with "no placeholders, so every copy would be identical" — about a merge that
would in fact have filled them.

### Fixed — three small format defects

- **`-0`.** `formatNumber` took the sign from the input rather than the rounded
  value, so `-0.4` at no decimal places printed `-0` — a quantity that does not
  exist, on a slide, from an ordinary cell.
- **`number:1e2` asked for a hundred decimal places and got them.** `Number`
  reads `1e2`, `0x10` and padded strings, so a spec that is not a count of
  places was used as one. Digits only now; anything else returns the cell as it
  stands, like every other unreadable format.
- **`date:MMMM` printed `MarM`** — `MMM` replaced and the fourth `M` left
  standing. `MMMM` is a full month name now rather than being refused, since it
  is a thing to want. The supported tokens are exactly the six the manual lists.

### Added — a placeholder in a chart is named instead of passed over

A chart's labels live in `ppt/charts/chartN.xml` and SmartArt's in
`ppt/diagrams/dataN.xml`, neither of which is a `<a:p>` on the slide. So
`fieldsIn` never saw them: the author put `{{Region}}` in a chart title, the
pane counted the placeholders it could see, and the braces shipped on every
merged slide. Not merging them is a stated limit; not saying so was the defect.

`prepareBlock` reads the parts each block slide RELATES to — never the package
at large, because below API 1.10 the template comes back as the whole deck and a
chart on slide 40 is not this block's problem. It uses `fieldsIn` on the parsed
part rather than a regex over the markup, because chart text is DrawingML and a
placeholder split across runs is the ordinary state of one after an edit.

A block whose only placeholders are in a chart used to be refused with "no
placeholders, so every copy would be identical" — true, and useless to somebody
looking at one. It now names them and says what to do.

### Fixed — a Danish date column formatted half of itself

`NAMED_DATE` admits `ÆØÅ`, so `detectType` types a Danish date column as a date.
The month NAME then went to `new Date`, which matches an English three-letter
prefix — so `marts` and `januar` worked, `maj` and `oktober` did not, and one
column came out half formatted and half raw across a merged deck with nothing
saying why.

The sharpest form of the rule this engine is built on: not a wrong date, but the
same column rendering two ways.

English, Danish, Norwegian and Swedish month names are a stated table now, full
and three-letter, listed rather than matched by prefix — an open prefix rule
reads `1 marketing 2026` as March. No word in the four means a different month
in another, and a test asserts that over the whole table rather than trusting
it, so a word added later that clashes fails there.

The `new Date` fallback stays, deliberately. It is what makes French and Italian
month names work for users the table does not list, and dropping it would turn a
partial answer into no answer for them. Its inconsistency is why the table
exists; it is a floor, not the mechanism.

**Not the Danish locale the backlog rejected.** That was a string table for the
pane's own text and is still rejected. This is reading the user's DATA, which
the date regex already reached for.

### Added — the run record is on screen while the run is still going

It was gated on the run being over, so a host that wedges produced a pane
showing "Waiting on PowerPoint…" and nothing else — on exactly the run somebody
needs to explain. The record is also SEEDED when the run starts rather than left
to the first trace that arrives after the subscriber, because the run's most
useful line, the environment, is written before it.

## [0.1.0] — 2026-08-27

The first release, and the first build that works at all: every one before it
never loaded Office.js, so the pane rendered its header and stopped, on every
host, with nothing said. Everything below shipped to `main` before this tag;
this is the point at which there is a manifest to download.

### Added — the package the engine hands over is checked as a package

Everything else in the suite tests a decision: does this paragraph merge, does
that plan skip the right row. None of it asked the question PowerPoint asks,
which is whether the file is a legal OOXML package — and that answer is binary
and expensive, because a deck that opens as "repaired" has lost whatever
PowerPoint decided to drop, in somebody's presentation.

`test/package-valid.test.ts` runs four real merges and checks the bytes: every
relationship resolves to a part that exists, no duplicate rIds, every part has a
content type, no override naming a part that is gone, slide ids unique and in
the format's range, and no slide part the deck does not list. Then it runs the
whole-deck route — keep what the run produced, drop the rest — and checks it
again, because that is where #38 lived and it is the only path that takes parts
OUT of a package.

**The content-type rule was toothless when first written, and injecting a defect
is what showed it.** A real .pptx declares `Default Extension="xml"`, so every
XML part passed. Deleting the clone's `addContentTypeOverride` — which would
ship every merged slide untyped — left the gate green. Slides and notes slides
are checked for an override of their own now, and both injected defects fail
four tests and one respectively.

### Fixed — a relationships path for a part at the package root

`Pkg.relsPathFor` used `lastIndexOf("/")` without handling -1, so a root part
answered `[Content_Types].xm/_rels/[Content_Types].xml.rels` — the last
character dropped and a directory invented. Nothing calls it that way today; it
would have failed silently when something did.

### Fixed — an impossible date was silently rolled forward

`2026-02-29` merged as **1 March** and `31 Feb 2026` as **3 March**, on every
slide, with nothing said. The manual promised the opposite and the engine's own
`utcDate` was written to enforce it.

The guard could never have fired on that path. `parseDate` handed the cell to
`new Date`, which NORMALISES, and then read the components back off the result —
by which time they were valid ones that round-tripped perfectly. It only ever
saw numbers something else had already made correct. The slash spellings took
their components from the string and were right all along, which is why the
manual's examples were all slash dates.

The two remaining spellings `looksLikeDate` admits now take their components
from the string too. A month NAME still needs the platform, and it is resolved
by parsing the first of that month — a day that exists in every month, so the
answer is the name's month and never a rollover. Asking `new Date` about the
whole cell is what let `31 Feb 2026` through as 3 March: the month had already
moved by the time anything looked at it, and the components then agreed with
themselves.

Found by running every accepting date form through the parser and diffing what
came back against what went in, rather than by reading it.

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
