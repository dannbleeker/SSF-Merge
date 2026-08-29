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

### One file per recipient
**Priority: medium.** Feasibility: medium.
Blocked by WebView2: blob downloads from a task pane do not work
([office-js#1511](https://github.com/OfficeDev/office-js/issues/1511)). The
route is Graph upload plus a link, or `openBrowserWindow` to a download page.

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
steps a chart already gets. **Filling** is not: a chartEx keeps its labels in
`<cx:pt>` elements this reader does not know, and its title in ordinary
DrawingML that would merge the moment the part were visited. Where the rest of
its text lives is a fact about a schema nobody here has a real file of.

So the blocker is a real .pptx containing one, not the code. Guessing the schema
from documentation is precisely what produced the SmartArt drawing bug, where
the fixture and the reader agreed with each other and disagreed with PowerPoint.
Build the file in PowerPoint first, the way the test kit's SmartArt was.

## Rejected — do not re-propose

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
