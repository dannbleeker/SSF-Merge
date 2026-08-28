# Backlog

The single curated list of what is open. Items graduate from here into a PR and
are **removed when they ship** — the README's feature table and the manual are
where shipped work is described, so anything still listed here is genuinely not
done.

Priority is what it costs the product to be without it, not how interesting it
is to build.

## Next

### Image fields — the aspect-ratio question
**Priority: high.** Feasibility: blocked, and blocked on a HUMAN rather than on
work. Probe question 8 — does `ShapeFill.setImage` stretch or preserve aspect
ratio — is the one no API reads back, so five sheets have not touched it and a
sixth will not either. Somebody has to fill a rectangle with an image whose
aspect ratio differs from the shape's and look at the slide. If it stretches,
the engine has to letterbox before sending, which is the whole design question.

The feature itself is in "After the first release" below; this is the single
measurement it waits on.

### What `exportAsBase64Presentation` drops, on a deck that can answer
**Priority: low.** Feasibility: high, and it needs one thing — a deck with
COMMENTS on it.

Asked twice now and unanswered twice for the same reason: both probe decks
carried no comments and no `ppt/authors.xml`, so there was nothing for the
export to drop. What it does leave behind is `ppt/webextensions/*` (which is
Script Lab's own registration, so the probe's deck will always show it),
`changesInfo1.xml` and `revisionInfo.xml` — none of them content, all rebuilt
by the host.

It matters only on the `subset` route (1.10), where the exported package is the
template block the clones are made from: comments on a template slide would be
dropped from every copy. Low because nobody has asked for comments to survive a
merge, and because the answer is one probe run away whenever a suitable deck
exists.

## After the first release

### Image fields
**Priority: high.** Feasibility: medium.
`{{Photo|image}}` filling a rectangle through `ShapeFill.setImage` (1.8), or a
picture cloned into the package with its media part. Blocked on probe question
**8** — the aspect-ratio one, which no API reads back, so somebody has to look
at a slide. If the host stretches, the engine has to letterbox before sending.
(The number has moved twice as questions were answered and renumbered; it is the
aspect-ratio one, whatever it is called this week.)

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

### Charts and SmartArt
**Priority: medium.** Feasibility: low to medium.
Text lives in `charts/chart*.xml` and `diagrams/data*.xml` with embedded
workbooks. Merging them is still open.

**The half that said "the pane must say so out loud rather than skipping a
field the user placed" SHIPPED on 2026-08-27.** `prepareBlock` reads the parts
each block slide relates to and reports what it finds there, step 2 names them,
and a block whose only placeholders are in a chart gets that sentence instead of
"no placeholders". So the silent case is closed and what is left is the feature
itself.

### A filter expression language
**Priority: low.** Feasibility: medium.
Row filters shipped as a searchable checkbox list. An expression — `Region =
"North" and Revenue > 1000` — is the thing a checkbox list cannot do on 400 rows,
and it is a parser, an evaluator and an error surface on a 340 px pane. Usage
will say whether it is worth answering; nothing has asked for it yet.

## Rejected — do not re-propose

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
