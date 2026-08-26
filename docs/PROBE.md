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

1. Open PowerPoint with a deck you do not mind four slides being added to and
   removed from. **Save it first anyway.**
2. Install [Script Lab](https://appsource.microsoft.com/product/office/wa104380862)
   if you do not have it: **Insert → Get Add-ins → search "Script Lab"**.
3. **Script Lab → Code → new snippet**, delete what is there, and paste the whole
   of [`probe/probe-snippet.ts`](../probe/probe-snippet.ts).
4. Press **Run**, then open the console pane at the bottom.
5. Copy everything between `=== SSF MERGE ANSWER SHEET ===` and `=== END ===`.

Then, in this repository:

```bash
npm run build:lib
node scripts/read-answers.mjs sheet.json --save
```

It prints what each answer means and files the sheet under
`docs/host-answers/`, stamped with when it was taken.

## What it asks

| # | Question | What turns on it |
| --- | --- | --- |
| 1 | Does a cloned slide with a **fresh creation id** insert cleanly? | The entire package path. If inserting fails, merged slides can only ever be delivered as a separate presentation |
| 2 | Does a tag written into the **package** read back through Office.js? | Undo, re-run, and every piece of merge metadata |
| 3 | Does `getSubstring(a, n).text = v` keep the run's formatting? | Whether live preview can be targeted, or has to redraw whole shapes |
| 4 | Do two writes queued in one batch interfere? | Whether replacements can be queued in any order or must go right to left |
| 5 | Does `fill.setImage` stretch or preserve aspect? | Image fields, and only those. **Nothing reads this back**, so it is a look-at-the-slide question for later |

Question 1 has **two arms**, and that is the point. One deck carries two slides
with different creation ids, which is what the engine produces; the other
carries the same two slides sharing one id, which is the shape
[office-js#6105](https://github.com/officedev/office-js/issues/6105) reports
failing. Asking only the first cannot tell "the bug is absent here" from "this
host refuses every insert".

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

## One answer is not evidence about your host

It is evidence about your host **in that minute**. A sibling project has
question-by-question records of the same build answering differently minutes
apart. If an answer decides something expensive, run the probe twice.
