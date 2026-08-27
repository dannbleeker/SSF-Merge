# Backlog

The single curated list of what is open. Items graduate from here into a PR and
are **removed when they ship** — the README's feature table and the manual are
where shipped work is described, so anything still listed here is genuinely not
done.

Priority is what it costs the product to be without it, not how interesting it
is to build.

## Next

### Picking the block by clicking slides
**Priority: medium.** Feasibility: unknown until a probe answers.
Two typed slide numbers work and are checked as they are entered, but clicking
the slides is what a user reaches for. `getSelectedSlides` is the obvious route
and this host is documented to wedge its whole selection subsystem after
`setSelectedShapes`, so it needs a probe question before it needs code — the
question being whether the SLIDE selection API is affected at all, or only the
shape one.

### The delimiter sniff reads only as far as the first newline
**Priority: low.** Feasibility: high.
`parseDelimited` decides tab-versus-comma from `src.slice(0, src.indexOf("\n") + 1)`.
A quoted FIRST header cell containing a newline puts that boundary inside the
cell, so the sample never reaches the tab and the whole table parses as one
column: `parseDelimited('"a\nb"\tc\nx\ty')` returns `[["a\nb\tc"], ["x\ty"]]`.

Found by an adversarial review of the pane controls and deliberately NOT fixed
there — it is engine code that commit did not touch, and the pane degrades
loudly rather than silently: one column means every placeholder is unmatched, so
`blockedReason` names them all and the merge button stays down. Narrow trigger
(the first header cell specifically, and only for tab-versus-comma), which is
why it is low rather than blocking.

The fix is a quote-aware sniff, or counting candidate delimiters across the
whole text and taking the majority.

## After the first release

### Image fields
**Priority: high.** Feasibility: medium.
`{{Photo|image}}` filling a rectangle through `ShapeFill.setImage` (1.8), or a
picture cloned into the package with its media part. Blocked on probe question
**5** — the aspect-ratio one, which no API reads back, so somebody has to look
at a slide. If the host stretches, the engine has to letterbox before sending.
(This said "question 4" until question 4 was answered and turned out to be the
offsets one.)

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

### Row filters
**Priority: medium.** Feasibility: high.
A checkbox list with search ships first. An expression language on a 340 px pane
is a v2 question, and usage will say whether it is worth answering.

### Charts and SmartArt
**Priority: medium.** Feasibility: low to medium.
Text lives in `charts/chart*.xml` and `diagrams/data*.xml` with embedded
workbooks. Until it is built, the pane must say so out loud rather than skipping
a field the user placed.

### Danish locale
**Priority: low.** Feasibility: high.
The string table exists from the first pane commit; the layout assumes nothing
about word length. English ships first because the Marketplace listing is English.

## Rejected — do not re-propose

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
  slowness: of 327 answered batches the slowest took 31s against a 45s budget,
  and seventeen abandoned calls never came back.

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
