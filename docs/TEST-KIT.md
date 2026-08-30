# SSF Merge — real-host test run

Everything shipped since v0.1.0 in one deck: text, formats, speaker notes,
picture fields, charts and SmartArt. It takes about ten minutes.

The point of this round is that the chart in the template was **not written by
SSF Merge**. It was written by a different tool, and the SmartArt on slide 3 by
PowerPoint itself. Every test in CI checks the engine against fixtures the
engine's own author wrote; this checks it against somebody else's file, which is
the one thing a test suite here cannot do.

That gap is not theoretical. On 2026-08-30 the integrity checker called a sound
package damaged the moment a real SmartArt was in it — sixteen problems, every
one of them PowerPoint's ordinary `r:blip=""` markup, which no fixture here had
ever written.

## Files

All of it lives in `test-kit/`.

| File | What it is |
| --- | --- |
| `SSF-Merge-test-template.pptx` | The template. Slide 1 is instructions, slides 2–3 are the block. Slide 3 carries the SmartArt, already authored. |
| `modern-chart.pptx` | A second, one-slide deck: a **sunburst** written by real PowerPoint. Its own short round, below. |
| `data.txt` | Three rows, tab-separated. Copy the whole thing, header included. |
| `ada.png`, `grace.png`, `alan.png` | The pictures. Deliberately one wide, one tall, one square. |

`test/test-kit.test.ts` merges this same template on every CI run, so the deck
you are asked to open cannot quietly stop working between rounds. What that test
cannot do is open it in PowerPoint, which is the whole reason this page exists.

**Driving it from Claude Code instead of by hand:** `test-kit/PROMPT.md` is the
prompt for that, with what to expect from Playwright against PowerPoint for the
web. Read its caveats first — the browser can drive the web host only, and the
merged file is the evidence either way.

`test-kit/driver/` holds the scripts that did it on 2026-08-28 and again on
2026-08-30: sideload the add-in, drive the pane, fetch the merged deck out of the
browser, and check the package. Its README names the traps, which are not
guessable, cost those rounds hours, and read as your own mistake when you hit
them — two of them make a working pane look broken. Start there rather than from
scratch.

## The SmartArt is already in the template

It used to be yours to add by hand at the start of every round. It is committed
now — authored in desktop PowerPoint on 2026-08-30, both halves checked before
use — so there is nothing to do here. It is still PowerPoint's own markup rather
than a fixture's, which is the whole reason the diagram is in this deck.

What that step guarded is worth keeping, because it will bite anyone who
re-authors it. A SmartArt stores its text **twice** — the model in `dataN.xml`
and the laid-out rendering in `drawingN.xml` — and PowerPoint writes the
rendering only when the diagram is done being edited. Save with the caret still
in a box and that box reaches the file with text in the model and **none in the
drawing**. It happened on 2026-08-28: `{{Name}}` was empty in `drawing1.xml`, so
every merged copy showed an empty first box and the merge took the blame for
what the template had done before it ever ran.

If you ever rebuild it: **Insert ▸ SmartArt ▸ Process ▸ Basic Process**, three
boxes reading `{{Name}}`, `{{Region}}` and `Renewal {{Renewal|date:d MMM}}`,
delete the grey box, **click empty space on the slide before saving**, then
confirm all three strings appear in `ppt/diagrams/drawing1.xml` and not only in
`data1.xml`.

And if a merged SmartArt ever looks blank, check the TEMPLATE's drawing part
before suspecting the merge.

## The run

1. Open the task pane.
2. **Step 1** — block **from 2 to 3**. It should say 2 slides.
3. **Step 2** — paste the contents of `data.txt`. It should read **3 rows** and
   name the columns. A picture picker appears underneath, saying **3 pictures
   named in Photo** — choose `ada.png`, `grace.png` and `alan.png`. It should
   then say **All 3 pictures matched.**
4. **Step 3** — the chips should list `Name`, `Region`, `Renewal`, `Revenue`,
   `Photo`, `Nickname`. `Nickname` should be marked as having no column.
5. **Step 4** — preview one row if you want to. Remove it afterwards.
6. **Step 5** — merge. It should add **6 slides**.

## What to check, in order of what would matter most

**1. The deck opens with no repair prompt.** Close PowerPoint after the merge,
reopen the file. This is the check the whole round is for: charts and SmartArt
add four to six parts per merged slide, and a package PowerPoint calls damaged
is repaired silently, dropping whatever it decided to drop.

**2. Each chart shows its own region.** Slides 1, 3 and 5 of the merged block
should read *Quarterly revenue — Nordics*, *— Benelux*, *— DACH*, and the
category axis under the first column should say the same. If all three show the
same region, the copies are sharing one chart.

**3. Right-click a chart ▸ Edit Data.** The workbook that opens should hold the
merged region, not `{{Region}}`. Close it. Then look at the chart again — it
must not have changed. This is the one an ordinary test cannot reach: PowerPoint
refreshes the chart from that workbook when Excel closes, so a merge that filled
only the chart would revert in front of you here.

**This check needs DESKTOP PowerPoint.** There is no Excel round-trip in
PowerPoint for the web, so a round run in the browser cannot answer it at all —
and "did not run" is not the same answer as "passed". If the round was driven in
the web host, say so and leave this one open, or open the merged deck in desktop
PowerPoint afterwards and do it there.

**4. The SmartArt shows that row's name.** Three merged copies, three different
names. SmartArt stores its text twice — a model and a rendering — and only the
rendering is on screen, so if the boxes read `{{Name}}` the wrong half was
filled.

**5. The pictures.** Each row gets its own photo, filling the orange frame with
no bands of background and no distortion. Two things on the photo tell you
which of the three possible outcomes you are looking at.

**The white border says which axis was cropped.** It runs around the edge of
each photo, so a correct cover-crop cuts the pair on the short axis and keeps
the pair on the long one:

| Photo | Shape | Border you should see |
| --- | --- | --- |
| `ada.png` | wide | **top and bottom only** — left and right cropped away |
| `grace.png` | tall | **left and right only** — top and bottom cropped away |
| `alan.png` | square | **left and right only** — the frame is wider than it is tall |

All four edges visible means it did not crop: either a stretch, or a letterbox
if there are bands of background as well.

**The four yellow dots say whether it distorted.** They sit inside the part of
each photo that a correct crop keeps, so **all four should be visible on all
three**, and each one should be **round**. An oval is a stretch, and nothing
else produces one.

> The dots used to sit at the photo's corners, and this page asked you to
> confirm they survived. They cannot: a correct crop removes a third of `ada`'s
> width and nearly two thirds of `grace`'s height, and the dots were within 5%
> of each corner — so every one of them went, on all three photos, no matter how
> well the merge worked. The round of 2026-08-28 duly read their absence as a
> failure. They were moved inward on 2026-08-29; if you are testing an older
> deck, expect no dots at all and judge by the border.

**6. `{{Nickname}}` is still visible** on each second slide. There is no such
column, and a placeholder with no column is meant to stay on the slide rather
than blank out — a blank slide looks finished when it is not.

**7. The formats.** `1 250 000 EUR`, and `Renewal 1 Mar 2026` / `15 Apr 2026` /
`30 May 2026`.

**8. Speaker notes.** View ▸ Notes. Each merged slide should read *Call Ada
before 1 Mar*, with the row's own name — the notes page is per-copy content.

**9. The pane's own line.** On 2026-08-30 it read, exactly:

> 6 slides added after slide 3 · 39 placeholders filled · 3 pictures placed.

Thirty-nine, not thirty — the older number in this file predated the chart and
SmartArt placeholders being counted, and a round comparing against it reported a
discrepancy that was not one. If it says **the data behind N charts could not be
merged**, that is the one soft failure in this list: the slides are still right
and only Edit Data is stale. Tell me the number.

**10. Undo.** Press it. Exactly the 6 merged slides should go, and the template
and slide 1 should stay.

## The second half of the run

These four had only ever been checked against fixtures until 2026-08-30, when
all four passed in PowerPoint for the web. They stay on the list because they
are cheap — the same `data.txt` and the same template do all of them, about
three minutes between them — and because three of them are about numbers the
pane PROMISES before it acts, which is exactly the kind of thing that drifts.

**1. The blank-cell control.** Paste `data.txt` as usual, then **clear one cell
in the paste box** — Grace's `Region` is the one to pick, since a chart, a
SmartArt and the slide text all name it. On the merge step, under the row list,
a line reads *A blank cell leaves a blank — change what happens*. Open it and
try all three:

| Choice | What the merged deck should show |
| --- | --- |
| **Leave the space empty** | Grace's slides have a gap where the region was — chart title, axis label, SmartArt box and body text |
| **Show the field, like `{{Region}}`** | those same four read `{{Region}}` |
| **Leave the whole row out** | Grace produces no slides at all: **4 slides added, not 6**, and the line above the button says why |

The third is the interesting one, because the number on the button has to change
the moment you choose it — the forecast and the plan are two different pieces of
code and they have disagreed before. In 2026-08-30 they agreed: the line became
*2 of 3 rows × 2 slides · 4 slides will be added after slide 3*, the button became **Add
4 slides**, and the merge then reported *4 slides added after slide 3 · 26
placeholders filled · 2 pictures placed · 1 row skipped for a blank field*.

> **If you are driving this from a script, use real key events.** It is a
> `<select>`, and setting `.value` with a dispatched `change` moves the DOM
> without moving React — the button goes on saying "Add 6 slides" and it looks
> exactly like the forecast bug this check exists to find. `cdp-key.mjs`.

**2. A semicolon-separated paste.** Take the same three rows, replace every tab
with `;`, and paste that instead. It should read **3 rows** and name the same
five columns. This is what Excel writes on any machine whose locale uses the
comma as a decimal point — Danish, German, French — so it is what a colleague's
export actually looks like, and it once produced one column named
`Name;Region;Revenue;Renewal;Photo` and a dead merge button.

Also try one with a decimal comma, `1250000` → `1,25`: that is the case that
made the first rule wrong. It reads as a DECIMAL separator, so `1,25`, `0,88`
and `1,64` render through `number:0` as `1 EUR`, `1 EUR` and `2 EUR`.

**3. What the pane says it did.** Send the whole summary line back whatever it
says — see check 9 above for what it said last time.

**4. Undo names its range.** Before pressing it, the card should name the slides
it is about to remove, and that range should match what step 5 said it added.
Last time: *"Remove slides 4 to 9, which this merge added"* against *"6 slides
added after slide 3"*, and *"Remove slides 4 to 7"* on the skip run.

Nothing here needs a re-install: the add-in is served from GitHub Pages and only
a change to the manifest itself would.

## What still has no real-host answer

**Edit Data, on both decks.** It needs DESKTOP PowerPoint, and both rounds so
far were driven in the browser. The bytes are right — each merged chart's
workbook holds its own region — but whether PowerPoint's Excel round-trip
preserves them has never been watched happen. "Did not run" and "passed" are
different answers.

**The preview step.** Step 4 is optional and has been skipped both times.

**The old-PowerPoint fallback notice.** Genuinely optional, and nothing from
2016 onwards will show it.

## The second deck: the modern chart

`modern-chart.pptx` is a separate thirty-second round, and it is worth doing
because it is a different kind of chart in a different part of the format.

A waterfall, funnel, treemap, sunburst, histogram, pareto, box-and-whisker or
region map is not a `<c:chartSpace>` at all — PowerPoint stores it as a chartEx
part under a Microsoft relationship, with its own way of keeping labels. The
chart in the main template is the classic kind, so nothing in that round touches
this code.

The deck holds one slide: a sunburst titled `{{Name}} pipeline`, a text box
reading `Pipeline for {{Name}}`, and a hierarchy whose outer ring's first cell is
`{{Region}}` with `Benelux` and `DACH` beside it as plain text.

`verify-package.mjs` checks this deck too now. It tells the two shapes apart by
which kind of chart part the deck holds, and answers **7/7** on a good sunburst
merge. It used to report 4/13 here — the main template's checklist run against a
deck that has no classic charts, no photos, no notes and no `{{Nickname}}` — and
the round of 2026-08-30 read that correctly as the wrong tool and checked the
sunburst by hand. A deck that is neither shape is now refused outright rather
than scored, because a red that means nothing is worse than no red.

1. Open it, block **from 1 to 1**, paste the same `data.txt`, merge. Three slides.
2. **Each chart's title names its own row** — *Ada pipeline*, *Grace pipeline*,
   *Alan pipeline*. Same title on all three means the copies share one chart.
3. **The outer ring reads that row's region, and only in the first segment.**
   Alan's should read `DACH`, `Benelux`, `DACH`: one merged, two untouched. All
   three segments changing is a merge writing where no placeholder was.
4. **The inner ring still reads `Existing` and `New`.**
5. **Right-click ▸ Edit Data**, close Excel, look again — as with the main deck,
   this is where a half-merge shows itself, and it needs desktop PowerPoint.
6. **Open it in an old PowerPoint if you have one to hand.** A modern chart
   carries a rendered picture for hosts that cannot draw it, and that picture is
   the TEMPLATE's data — another recipient's figures under this recipient's
   name. Each merged copy should show a bordered notice saying the chart needs a
   newer PowerPoint instead. Nothing from 2016 onwards will ever show it, so
   this one is genuinely optional.

The deck was recorded on a machine with think-cell installed, which stamps every
deck it touches. `test-kit/strip-thinkcell.py` is the one edit made to the file
before it was committed, and says exactly what it removed.

## What to send back

- Whether it opened clean, and the repair prompt's wording if it did not.
- The merged deck itself, if anything looks wrong — the file says more than a
  description of it.
- The pane's summary line.
- Anything on the list above that did not match.
- Which host you drove, and what you therefore did NOT run.

Save the deck to `test-kit/out/round-<date>.pptx` and run the checker over it:

```bash
node test-kit/driver/verify-package.mjs test-kit/out/round-<date>.pptx
```

It should say **13/13** for the main deck and **7/7** for the sunburst, and it understands a deck that still holds its template
block — it counts the parts the MERGED slides reach, not the parts in the
package. If you intend to believe its verdict, run `mutate.mjs` over the same
deck first and read the line that says how many guards it actually proved.

If the deck fails to open at all, that is still a good result: send the file and
the message. It is the failure mode this round exists to find, and the merged
deck plus the template together are enough to locate it.

**Two habits worth keeping, both learned the expensive way.** A check that fires
on PowerPoint's own markup is a bug in the check, not in the merge — confirm
what the tool is objecting to before writing it up. And a tool that reports
success without saying what it proved has told you nothing: read the counts.
