# SSF Merge — real-host test run

Everything shipped since v0.1.0 in one deck: text, formats, speaker notes,
picture fields, charts and SmartArt. It takes about ten minutes.

The point of this round is that the chart in the template was **not written by
SSF Merge**. It was written by a different tool, and the SmartArt will be
written by PowerPoint itself. Every test in CI checks the engine against
fixtures the engine's own author wrote; this checks it against somebody else's
file, which is the one thing a test suite here cannot do.

## Files

All of it lives in `test-kit/`.

| File | What it is |
| --- | --- |
| `SSF-Merge-test-template.pptx` | The template. Slide 1 is instructions, slides 2–3 are the block. |
| `data.txt` | Three rows, tab-separated. Copy the whole thing, header included. |
| `ada.png`, `grace.png`, `alan.png` | The pictures. Deliberately one wide, one tall, one square. |

`test/test-kit.test.ts` merges this same template on every CI run, so the deck
you are asked to open cannot quietly stop working between rounds. What that test
cannot do is open it in PowerPoint, which is the whole reason this page exists.

**Driving it from Claude Code instead of by hand:** `test-kit/PROMPT.md` is the
prompt for that, with what to expect from Playwright against PowerPoint for the
web. Read its caveats first — the browser can drive the web host only, and the
merged file is the evidence either way.

`test-kit/driver/` holds the scripts that did it on 2026-08-28: sideload the
add-in, drive the pane, fetch the merged deck out of the browser, and check the
package. Its README names the traps, which are not guessable and cost that round
hours. Start there rather than from scratch.

## Before you start: add the SmartArt

Nothing outside PowerPoint can author a SmartArt graphic, so this one step is
yours — and a diagram PowerPoint wrote is a better test than one I wrote.

1. Open `SSF-Merge-test-template.pptx`, go to **slide 3**.
2. **Insert ▸ SmartArt ▸ Process ▸ Basic Process.**
3. Type into the three boxes:
   - `{{Name}}`
   - `{{Region}}`
   - `Renewal {{Renewal|date:d MMM}}`
4. Delete the grey instruction box.
5. **Click somewhere else on the slide before saving**, so the last box you
   typed into loses focus.
6. Save.

Step 5 is not fussiness. A SmartArt stores its text twice — the model in
`dataN.xml` and the laid-out rendering in `drawingN.xml` — and PowerPoint writes
the rendering when the diagram is done being edited. Save with the caret still
in a box and that box can reach the file with text in the model and **none in
the drawing**. It happened on 2026-08-28: the `{{Name}}` box was empty in
`drawing1.xml`, so every merged copy showed an empty first box, and the merge
took the blame for something the template did before it ever ran.

Worth checking if anything looks blank later: unzip the template and confirm all
three strings appear in `ppt/diagrams/drawing1.xml`, not just in `data1.xml`.

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

**9. The pane's own line.** It should read something like *6 slides added after
slide 3 · 30 placeholders filled*. If it says **the data behind N charts could
not be merged**, that is the one soft failure in this list: the slides are still
right and only Edit Data is stale. Tell me the number.

**10. Undo.** Press it. Exactly the 6 merged slides should go, and the template
and slide 1 should stay.

## What to send back

- Whether it opened clean, and the repair prompt's wording if it did not.
- The merged deck itself, if anything looks wrong — the file says more than a
  description of it.
- The pane's summary line.
- Anything on the list above that did not match.

If the deck fails to open at all, that is still a good result: send the file and
the message. It is the failure mode this round exists to find, and the merged
deck plus the template together are enough to locate it.
