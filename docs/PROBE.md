# The host probe

The questions that only a real PowerPoint can answer, and one at the end that
only a person can — the table under *What it asks* is the whole list. Nothing in
the add-in should be built on a guess about any of them.

Running it takes about two minutes and leaves your deck as it found it.

## Why it exists

The whole design rests on assumptions about a host that is documented, in
`CLAUDE.md`, to lie about object ids, to accept calls it does not perform, and
to answer differently on two runs of the same build. A sibling project spent
months learning that by inference. This probe asks directly instead.

## Running it

**Work on a throwaway copy of a deck, not a real one.** The probe adds four
slides and removes them again, and on the web that is not undoable in the way
you would expect: PowerPoint for the web has AutoSave permanently on, so every
step is written to OneDrive as it happens. "Save it first" is desktop advice and
does not protect anything here. Any deck with at least one slide will do.

### PowerPoint for the web

1. Open your throwaway deck in PowerPoint for the web, in the **editor**, not the
   read-only viewer. Script Lab cannot load in the viewer.
2. **Home → Add-ins** (or **Insert → Add-ins**), search for
   [Script Lab](https://appsource.microsoft.com/product/office/wa104380862),
   and **Add** it. A **Script Lab** tab appears on the ribbon.
3. **Script Lab → Code**. The editor opens in the task pane on the right,
   showing whatever snippet was last open.
4. Open the pane's menu (top left of the editor) and choose **New Snippet**, so
   nothing of a previous one is left behind.
5. Make sure the **Script** tab is selected — not HTML, CSS or Libraries —
   select everything in it, and replace it with the whole of
   [`probe/probe-snippet.ts`](../probe/probe-snippet.ts). It is about 24 KB; the
   [raw file](https://raw.githubusercontent.com/dannbleeker/SSF-Merge/main/probe/probe-snippet.ts)
   is the easiest thing to select-all and copy. Leave the other three tabs alone;
   the snippet needs no libraries beyond the Office.js a blank snippet already
   carries.
6. **Script Lab → Run**. The pane switches to the runner. It takes under a
   minute; you will see slides appear at the end of the deck and disappear again.
7. Expand the **console** strip at the bottom of the runner pane.
8. Copy everything between `=== SSF MERGE ANSWER SHEET ===` and `=== END ===`.

If the console shows `the probe itself failed:` instead of an answer sheet, send
that line too — it is an answer about the host as well.

### Desktop

Same steps, with **Insert → Get Add-ins** to install and the runner opening in
its own pane. Desktop is the friendlier host and the less interesting one: the
web is where the ids are refused and the collections come back empty, so run the
web first.

Then, in this repository:

```bash
npm run build:lib
node scripts/read-answers.mjs sheet.json --save
```

It prints what each answer means and files the sheet under
`docs/host-answers/`, stamped with when it was taken.

**A question the probe could not put is reported as unknown, never as no.** The
tag read lands on the last slide in the deck, which is the inserted one only if
the insert worked; when it did not, that read falls on a slide the user owns and
has never carried our tag. The first sheet reported "the metadata scheme needs
rethinking" on exactly that read, about a scheme that had not been tested at
all. `tagVerdict` guards it now.

## What it asks

| # | Question | What turns on it |
| --- | --- | --- |
| 0 | **Control:** does this host insert the deck it just saved itself? | Whose fault a refusal is. Without it, `InvalidArgument` is equally our package and this host, and those are opposite conclusions |
| 1 | Does a cloned slide with a **fresh creation id** insert cleanly? | The entire package path. If inserting fails, merged slides can only ever be delivered as a separate presentation |
| 2 | Does a tag written into the **package** read back through Office.js? | Undo, re-run, and every piece of merge metadata |
| 3 | Does `getSubstring(a, n).text = v` keep the run's formatting? | Whether live preview can be targeted, or has to redraw whole shapes |
| 4 | Do two writes queued in one batch interfere? | Whether replacements can be queued in any order or must go right to left |
| 5 | Does `exportAsBase64Presentation` drop parts the file route keeps? | The template is READ through that export. A part it drops is a part every merged slide is missing, in a file that opens cleanly |
| 6 | Does a collection load of the deck's slides answer in full? | `deckSlideIds`, and through it every block lookup. A short read past the ceiling refuses a block; a scrambled one clones slides nobody chose |
| 7 | Does a slide insert survive a **standing selection**? | Whether the preview step must call `setSelectedShapes` — the one call in this family with a measured history of wedging the host |
| 8 | Does `fill.setImage` stretch or preserve aspect? | Image fields, and only those. **Nothing reads this back**, so it is a look-at-the-slide question for later |

Question 1 has **four arms**, and that is the point. One deck carries two slides
with different creation ids, which is what the engine produces; the other
carries the same two slides sharing one id, which is the shape
[office-js#6105](https://github.com/officedev/office-js/issues/6105) reports
failing. Asking only the first cannot tell "the bug is absent here" from "this
host refuses every insert".

The first real sheet made that concrete, and needed two more arms. Every insert
came back `InvalidArgument`, which admits two readings pointing opposite ways:
this repo generates a package PowerPoint will not take, or this host takes no
package at all. One is a morning's work; the other ends the package path. So:

- a **control arm** inserts the bytes of the presentation the probe is running
  in, read back through `getFileAsync`. That deck is a package PowerPoint wrote
  seconds earlier, so it cannot be malformed — a host that refuses it is
  refusing insertion itself. It runs first, while the deck is still only the
  user's own;
- a **theme arm** inserts the same package under `UseDestinationTheme` instead
  of `KeepSourceFormatting`. Only the latter has to import the source theme, so
  the pair separates a theme this host will not read from a package it will not
  read.

`insertionBlame` is the reading, and it says CANNOT TELL rather than guessing
when the control did not run.

### Does the export drop parts the file route keeps?

This add-in reads its template through `exportAsBase64Presentation` on any host
that has it, and then clones what comes back. office-js#6867 reports the
slide-level `exportAsBase64` omitting modern comments and `ppt/authors.xml`.
Nobody has checked whether the presentation-level call does the same — and if it
does, every merged slide is missing those parts, silently, in a file that opens
cleanly.

The arm exports **every slide in the deck** and compares its part list against
the same deck read through `getFileAsync`, which returns the package unfiltered.
Both arms cover the same slides, or the export would legitimately lack the parts
belonging to slides it was not asked for.

Part NAMES only, never content: a part list is structure, and this sheet is
written to be pasted into an issue where the merged values would be somebody's
salary review.

**It needs a deck with comments in it.** A deck that has none cannot answer the
question, and `exportPartsVerdict` says NOT ASKED rather than reporting the
absence as "the export keeps everything" — which would be the same mistake as
reading a question the run never put as a `no`. Run this one on a deck somebody
has actually commented on.

## What it does to your deck

It appends four slides, asks its questions, and removes them again by
**position**, never by id — a slide a run has just added does not round-trip
through `slides.getItem(id)` on the web, and a sibling project's by-id clean-up
once reported 45 successful deletes having removed nothing.

The sweep is clamped so it can never remove more than the probe added, never
more than the deck grew, and never reach an index below the deck's size when it
started. Those clamps are tested, and each one is proven to be load-bearing by
removing it and watching the tests go red.

If the deck ends larger than it started, the reader says so rather than
pretending it is clean.

## What it has answered so far

Six sheets as of 2026-09-01, all PowerPoint for the web, filed under
`docs/host-answers/` — that directory is the count, not this line.

The first answered nothing about the host and three things about the probe: it
gave a verdict on a question it had not asked, it named none of the four calls
that could have thrown, and its fixture deck was malformed. The second, once the
fixture was fixed and the control arm added, landed every insert and read the
package tag back.

The third answered the last two. A targeted substring write keeps the formatting
around it, so live preview can target a substring rather than redrawing whole
shapes. And two writes queued in one batch **do** interfere — the second sees the
first one's result — so Office.js replacements must be queued right to left.

It also caught the reader believing an error over a measurement: an insert timed
out and had landed both its slides anyway, and reading the error as decisive
produced three false statements in one run. The delta is the evidence.

The fourth added three arms and answered NONE of them, which is the honest
outcome rather than a wasted round:

- **the export** (question 5) ran on a deck with no comments and no
  `ppt/authors.xml`, so there was nothing for it to drop. It did leave five
  parts behind — the two web-extension parts and their rels, `changesInfo1.xml`
  and `revisionInfo.xml` — none of which is content, all of which the host
  rebuilds. That is worth knowing and it is not the question. **Re-run it on a
  deck carrying comments.**
- **the deck read** (question 6) answered all 8 of 8, in order. The ceiling
  [office-js#4272](https://github.com/officedev/office-js/issues/4272) describes
  is around fifty, so an eight-slide deck says the collection is not broken
  outright and says nothing about the case `ID_PAGE = 20` exists for. **Re-run
  it on a deck of more than fifty slides.**
- **the standing selection** (question 7) ran with nothing selected. The arm is
  read-only about the selection by design — the workaround it would justify is
  `setSelectedShapes` — so it can only observe a condition the user made.
  **Re-run it with a shape clicked.**

Each of those reports `unknown` with the re-run named. A question the run could
not put is never recorded as an answer.

**The fifth sheet (2026-08-28) is that re-run, and it answered two of the
three.** It was taken on a 26-slide deck with three shapes selected, which is
what made the difference — the arms were fine, the DECK was too small and too
idle.

- **the deck read** (question 6) answered all **58 of 58**, in order, nothing
  short and nothing empty. Fifty-eight is above the ceiling #4272 describes, and
  `canAnswerFiftyQuestion` says so, so this is an answer rather than a deck that
  could not reach the question. office-js#6363 does not reproduce here either.
  `deckSlideIds` keeps paging anyway — see `CLAUDE.md` for why that is a
  decision and not an omission.
- **the standing selection** (question 7) ran with **three shapes selected** and
  both slides landed. A slide insert is not the shape case #2775 and #3698
  describe, and nothing here needs `setSelectedShapes`.
- **the export** (question 5) was still unanswered here, for the same reason as
  before: the deck carried no comments and no `ppt/authors.xml`, so there was
  nothing to drop.

**The sixth sheet answered it, on a deck with four comments on it.** Four comment
parts and `ppt/authors.xml` went in and **none came out** — so office-js#6867
reaches the presentation-level call too. What that exposed is in `CLAUDE.md`: the
two template routes disagreed about comments, and `cloneSlide` was the half that
was wrong.

Read that sheet for question 5 alone. It was taken on a four-slide deck with
nothing selected, so its deck-read and standing-selection arms could not put
their questions — the reader says `unknown` for both, and the fifth sheet's
answers stand. **A later sheet does not overwrite an earlier answer it could not
ask.**

**Question 8 is answered too, on 2026-08-28, and it is the only one whose
evidence is a screenshot.** `ShapeFill.setImage` **STRETCHES**: three rectangles
filled from one square card — 1:1, 2:1 wide, 1:2 tall — put a round circle in
the square box, a wide ellipse in the wide one and a tall ellipse in the tall
one, with all four corner labels surviving in every box. Corners intact rules
out a crop; no bars rules out a letterbox.

`scripts/build-aspect-probe.mjs` rebuilds that instrument. It emits
`probe/aspect-probe.ts` and the card it fills with, and it is worth keeping
because the answer is a property of the HOST and hosts change.

So every question this probe was built to ask is answered, and the package path
is measured rather than assumed.

## One answer is not evidence about your host

It is evidence about your host **in that minute**. A sibling project has
question-by-question records of the same build answering differently minutes
apart. If an answer decides something expensive, run the probe twice.
