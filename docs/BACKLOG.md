# Backlog

The single curated list of what is open. Items graduate from here into a PR and
are **removed when they ship** — the README's feature table and the manual are
where shipped work is described, so anything still listed here is genuinely not
done.

Priority is what it costs the product to be without it, not how interesting it
is to build.

## Next

**Nothing is waiting on an answer any more.** This section held the probe
questions, and every one it was built to ask has been put and answered — the
last of them, the aspect-ratio one, on 2026-08-28. `docs/PROBE.md` has what they
said.

The first release is out, so the section below is not "later" any longer: it is
the work, in priority order.

## After the first release

### Excel via Microsoft Graph
**Priority: high.** Feasibility: medium.
A named table on OneDrive or SharePoint, read through `/workbook/tables/{id}/rows`.
Refreshable, shareable, and the reason people ask for this. Needs nested app
authentication, which removes a middle tier rather than adding one.

### Modern chart types
**Priority: medium.** Feasibility: medium.
A waterfall, funnel, treemap, sunburst, histogram or box-and-whisker chart is
not a `<c:chartSpace>`. PowerPoint stores it as a separate part reached through
`…/2014/relationships/chartEx`, and nothing here knows that relationship type.

Measured rather than assumed, and it is worse than "not filled". Merging a deck
carrying one produces ONE chartEx part shared by the template slide and every
merged copy, still holding `{{Region}}` and `{{Name}}` — which is exactly the
failure `docs/TEST-KIT.md` check 2 exists to catch. `test/chart-modern.test.ts`
pins it, and starts failing the day somebody adds support.

Two halves, and they are not equally safe. **Cloning** is schema-independent —
copy the part, repoint the relationship, add the content type, the same three
steps a chart already gets. **Filling** is not, and the research below is why.

#### Read out of real files, not reconstructed

Three chartEx parts, on 2026-08-29, out of LibreOffice's own chart test data —
real Office output. Not committed here: MPL-2.0 test data, and it is the
findings that are worth keeping. Re-fetchable from
[LibreOffice/core](https://github.com/LibreOffice/core) at
`chart2/qa/extras/data/xlsx/waterfall2.xlsx`, `.../funnel1.xlsx` and — the one
that matters, a chart on a SLIDE with its embedded workbook and its fallback
picture — `chart2/qa/extras/data/pptx/funnel-pp1.pptx`.

**A modern chart keeps its text in four places, and one picture of it.**

| Where | What | Merged how |
| --- | --- | --- |
| `cx:strDim > cx:lvl > cx:pt` | category labels | plain text node |
| `cx:tx > cx:txData > cx:v` | series names | plain text node |
| `cx:txPr > a:p > a:r > a:t` | title, axis titles | DrawingML |
| the embedded workbook's `sharedStrings.xml` | every one of the above, again | as a chart's already is |
| `mc:Fallback > p:pic` | a rendered PNG of the whole chart | **not at all** |

Four traps, each of which would have been got wrong by writing the reader first.

**`<cx:pt>` is the same element for text and for numbers.** `cx:strDim` holds
the labels; `cx:numDim` holds the values a chart plots, in `<cx:pt>` too. Scope
by the DIM, never by the element name — the identical mistake `<c:v>` in
`strCache` versus `numCache` already forced on the classic path.

**`<a:t>` is sometimes a cached copy and sometimes the only copy.** `waterfall2`
writes its title as `<cx:tx><cx:txData><cx:v>` AND again as DrawingML inside
`<cx:txPr>`; `funnel1` and `funnel-pp1` have no `<cx:tx>` at all and keep the
text only in the `txPr` run. `mergeDocument` walks every `<a:p>` in a part, so
the moment a chartEx is visited it fills the DrawingML and nothing else — one
shape merges by luck, the other keeps displaying the placeholder from `<cx:v>`
while the file carries the merged string in its formatting run. Text kept twice
and the engine filling the half nobody looks at is the SmartArt
model-and-drawing defect for the third time. Fill both, and test both shapes:
one file each is not a schema.

**The category labels are repeated once per series.** `funnel-pp1` carries three
`<cx:data>` blocks, each with its own copy of `Thing 1…4`. All of them move
together or the chart disagrees with itself.

**The fallback picture cannot be merged, ever.** It is a PNG of the chart as the
template drew it, shown by any host that cannot read chartEx. A merged copy
cannot regenerate it: this project has no chart renderer and is not going to
grow one for eight layouts it does not draw.

What the ecosystem does with that branch, looked up rather than reasoned about,
splits on exactly one thing — whether the tool can render:

| Tool | What it writes in `mc:Fallback` |
| --- | --- |
| PowerPoint | `p:pic`, a rendered PNG of the chart. It has a renderer and regenerates on save |
| LibreOffice, exporting a chartex | a plain rectangle: white fill, thin outline, `noTextEdit`, carrying the standard *"This chart isn't available in your version of…"* sentence. No picture (`writeChartexAlternateContent` in `oox/source/export/chartexport.cxx`) |
| python-pptx | nothing: `mc:AlternateContent` is invisible to it ([#621](https://github.com/scanny/python-pptx/issues/621)), so whatever was there rides along untouched and stale |

`mc:Fallback` is `minOccurs="0"` in the MCE schema, so a bare `mc:Choice` is
legal too. That gives three options, and SSF Merge is in LibreOffice's position
rather than PowerPoint's.

**Recommended: write the explanatory shape, LibreOffice's answer. NOT DECIDED —
the owner's call.**

Keeping the template's picture is the one to refuse. On a mail merge it does not
merely show stale data, it shows ANOTHER RECIPIENT'S figures under this
recipient's name, on any host that reads the fallback — a confidentiality
problem rather than a cosmetic one, and the merged deck is the artefact that
gets sent out. It is also the same mistake this engine already refuses
everywhere else: a placeholder with no column stays visible rather than
blanking, precisely because something that looks finished and is not is worse
than something that admits it.

Dropping the branch is honest and silent: an old host shows a hole and no reason
for it. The explanatory shape costs the same to write, says why, and has the
merit of being what a major producer already ships — so the shape is known-good
to PowerPoint. Its wording should name PowerPoint rather than Excel, which is a
bug in the sentence LibreOffice copied.

Two things that bound the decision. Nothing on a modern host ever sees this
branch — it takes `mc:Choice` — so the cost of getting it wrong falls only on
PowerPoint 2013 and earlier and on third-party viewers. And replacing the
picture lets the merged deck drop N stale PNGs, which is smaller as well as
honester.

And one piece of good news, worth as much as the traps: **the embedded workbook
hangs off the chartEx under the ORDINARY
`…/officeDocument/2006/relationships/package` type**, with
`<cx:externalData r:id="…">` naming it from inside `cx:chartData`. So
`chartWorkbooksOf` and `cloneChartWorkbook` work on it unchanged; what is
missing is only that `graphicPartsOf` does not recognise the chartEx
relationship in the first place. The chart's own `chartStyle` and
`chartColorStyle` parts are read-only and should stay shared, exactly as a
classic chart's are.

**What the engine does with one today, measured rather than predicted.** A real
`funnel-pp1.pptx` with `{{Region}}` in a category label, `{{Name}}` in a series
name and in the title, merged for two rows:

    fields reported        ["Name"]        — only the slide's own text box
    chartEx parts          1               — shared by the template and both copies
    embedded workbooks     1               — likewise
    every placeholder in the chart          survived

So it is worse than "not filled": the pane never SEES those fields. The step
that lists what is on the slides would not offer a Region column, and a block
whose only placeholder is in the chart is refused as having none.

#### The slide-side wrapper, confirmed

    <mc:AlternateContent>
      <mc:Choice Requires="cx2" xmlns:cx2="…/office/drawing/2015/10/21/chartex">
        <p:graphicFrame>…<a:graphicData uri="…/office/drawing/2014/chartex">
          <cx:chart r:id="rId2"/>
      <mc:Fallback>
        <p:pic>…<a:blip r:embed="rId3"/>   → ppt/media/image1.png

Both branches carry the same shape id and the same `a16:creationId`. The
`Requires` token is **cx2** here and is a 2015 namespace, so a reader keying on
`cx1` would miss this file entirely: key on the RELATIONSHIP type, which is
stable, not on the compatibility token.

#### What is left

Not a blocker any more, an implementation with a decision in it. Recognise the
chartEx relationship in `graphicPartsOf` and `cloneSlideGraphics`; fill the four
places above; scope `<cx:pt>` by its dim; decide what a merged copy does about
the fallback picture. The kit should grow a chartEx slide at the same time, and
that one still wants a PowerPoint-authored file — the same thirty seconds that
got it its SmartArt.

## Rejected — do not re-propose

- **Somewhere other than the current deck to put a merge.** Dropped by the owner
  on 2026-08-29, covering both destinations the manual had listed as designed:
  **one file per recipient**, and **into a new presentation**.

  Neither is a bad idea and neither is expensive to want. One file per recipient
  is what you attach to an email, and it is the only version of this that keeps
  a recipient from scrolling to a competitor's page. A new presentation avoids
  the deck that is slow to edit. They were weighed together because they are one
  plumbing problem in two sizes: a destination that is not the presentation the
  add-in is running in.

  What settled it is that the two halves are not equally cheap, and the cheap
  half buys the least. The per-recipient half is genuinely BLOCKED — a task pane
  cannot hand you a file
  ([office-js#1511](https://github.com/OfficeDev/office-js/issues/1511)), so it
  needs a Graph upload and a link, which means the nested app authentication
  Excel-via-Graph needs anyway. The new-presentation half is not blocked at all:
  `PowerPoint.createPresentation(base64)` is PowerPointApi 1.1, below this
  add-in's floor, and the pipeline already ends at the exact base64 string it
  wants — one call swapped, a day's work at most, and the merged deck inherits
  the master, layouts and theme for free because the package is a clone of the
  user's own.

  The cost that is easy to miss is verification. The add-in runs in the ORIGINAL
  presentation and cannot see the new one, so `createPresentation` answers
  `Promise<void>` and nothing else. Every other path here proves what landed by
  re-counting the deck, precisely because this host accepts calls it does not
  perform — and that evidence is simply unavailable through this door. A merge
  that cannot say whether it worked is a poor trade for a deck that opens in a
  second window.

  Reviving it means somebody actually sending merged decks out one per
  recipient, which is the blocked half — so the thing to watch for is that
  request, not the easy half's availability. If Excel-via-Graph ships, the
  authentication this needs already exists and the calculation changes.

- **A filter expression language.** Dropped by the owner on 2026-08-29. Row
  filters shipped as a searchable checkbox list and nothing has asked for more.
  The entry stood on a guess about what users would eventually want, which is
  the weakest reason to carry work: an expression like `Region = "North" and
  Revenue > 1000` is a parser, an evaluator and an error surface, on a 340 px
  pane, for a problem no one has reported having.

  Reviving it means somebody hitting the wall the checkbox list actually has —
  a filter too long to tick on a few hundred rows — and saying so. That is a
  fact about usage rather than a prediction, and it is what this entry was
  waiting for and never got.

- **A Danish locale.** Dropped by the owner on 2026-08-27. The entry that
  carried it was wrong about the cost in the direction that matters: it claimed
  a string table "exists from the first pane commit" and there is none. Every
  user-visible string in `steps.ts`, `render.ts` and `summary.ts` is an inline
  literal, and a good number are assembled from fragments — counts,
  pluralisation, "slide" versus "slides". So the work is not translation, it is
  building the table first, and the assembled sentences do not survive
  extraction unchanged: Danish pluralises and inflects differently, so anything
  built by concatenating a count onto an English noun has to become a
  whole-sentence lookup before it can be translated at all. The Marketplace
  listing is English and nothing has asked for Danish. Reviving this means
  reviving the string-table work, which is the larger half.

  **This is about the pane's OWN TEXT, and it does not cover reading Danish
  DATA.** Confirmed by the owner on 2026-08-27, when the month-name table
  shipped: `looksLikeDate` has admitted `ÆØÅ` since it was written, so a Danish
  date column was always going to reach the parser, and until that table it was
  read by accident — `new Date` matches an English three-letter prefix, so
  `marts` worked and `maj` did not, in the same column. Fixing that is
  correctness, not localisation.

  The distinction to carry forward: **what the user's SPREADSHEET says is data
  and gets read properly; what SSF Merge says back is English.** A merged deck's
  month names come out English too, and changing that is the string-table work
  above rather than an extension of the table.

- **A preview that writes onto the template slide and restores it.** This was
  the specified design and it is refused for the reason immediately below:
  setting a shape's text through Office.js re-authors it, and RESTORING goes
  through the same API that did the damage, so the text returns and the
  formatting does not — silently, on the master copy every merged slide is
  cloned from. A preview is an ordinary one-row merge, inserted and then swept
  by the same clamped positional sweep an undo uses. The `SSF_MERGE_TEMPLATE`
  tag that existed only to hold the text to put back is gone with it.
- **Merging through the PowerPoint API, shape by shape.** Setting text
  re-authors it (office-js#5858), and a sibling add-in measured a 680-second run
  that shipped duplicate slides. The whole architecture is the answer to this.
- **Any product name starting with `Power`.** Screened: Microsoft's trademark
  family is dense in exactly this space, and the name would never be ownable.
- **Waiting after `slides.add()`.** Tried on the sibling project; it cost 18 of
  19 probe answers in one round. This host is not the one the issue describes.
- **Raising the batch timeout to survive a stall.** A stall is death, not
  slowness. On the sibling project the batches that ANSWERED topped out at
  31.1s against a 45s budget, nothing has ever landed in the band between, and
  no abandoned call has ever come back late.

  Not contradicted by the probe's insert budget going 30s to 60s, though it
  looks like it. That call ANSWERED — the deck delta showed both slides had
  landed while the wait was still running — so the budget was short for work
  that succeeded, which is a different thing from waiting longer on a call that
  is already dead. Read `BUDGET` in `src/host/timeout.ts`: the numbers differ
  per call because the evidence does.
- **Bindings as a way round the id refusals.** Asked and answered on the sibling
  project: the host rejects the batch carrying the binding, with a control arm
  proving it was the binding.
- **Detecting charts whose tags were lost to cut/paste.** The count would be
  swamped by the tag writes this host refuses anyway, and would report "this
  host is unwell" rather than "your paste broke a slide".
