# The host probe

Four questions that only a real PowerPoint can answer, and one that only a
person can. Nothing in the add-in should be built on a guess about any of them.

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
| 5 | Does `fill.setImage` stretch or preserve aspect? | Image fields, and only those. **Nothing reads this back**, so it is a look-at-the-slide question for later |

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

Two sheets, both PowerPoint for the web, 2026-08-26, filed under
`docs/host-answers/`.

The first answered nothing about the host and three things about the probe: it
gave a verdict on a question it had not asked, it named none of the four calls
that could have thrown, and its fixture deck was malformed. The second, once the
fixture was fixed and the control arm added, landed every insert and read the
package tag back.

So the package path is measured, not assumed: cloned slides insert, and merge
metadata written into the file survives into the host's object model. What is
still open is question three and four, which the second sheet could not reach —
a shape proxy does not survive a `context.sync()` on this host, so the
experiments are queued in one batch now.

## One answer is not evidence about your host

It is evidence about your host **in that minute**. A sibling project has
question-by-question records of the same build answering differently minutes
apart. If an answer decides something expensive, run the probe twice.
