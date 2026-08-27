# Backlog

The single curated list of what is open. Items graduate from here into a PR and
are **removed when they ship** — the README's feature table and the manual are
where shipped work is described, so anything still listed here is genuinely not
done.

Priority is what it costs the product to be without it, not how interesting it
is to build.

## Next

### What the deck-read probe answers, and what to do about it
**Priority: medium.** Feasibility: high once the sheet arrives.
`deckSlideIds` already pages by position, so office-js#4272 cannot bite the way
it would have — but the probe's `deckRead` block asks three things nobody here
has measured, and each has a next step:

- `short` — whether a collection load answers short on this host at all. If it
  never does, the paging is insurance and can stay as insurance.
- `prefixOk` — whether a short read is the first n IN DECK ORDER. This is the
  one that decides severity for anyone who reaches for `load("items/id")`
  again: prefix-stable means a wrong block is REFUSED, scrambled means
  `indexOf` answers the wrong slide NUMBER and a merge clones slides nobody
  chose. If it comes back false, that fact belongs in `CLAUDE.md` as a
  never-do.
- `empty` — office-js#6363, a read that returns nothing after a sync that
  succeeded. If it reproduces, `blockIds` and `blockFromSelection` should say
  "PowerPoint would not list the deck" rather than "the deck has 0 slides",
  which is what they say today.

`canAnswerFiftyQuestion` reports whether the deck was big enough to answer the
first one at all. A nine-slide deck cannot, and a sheet from one must not be
read as though it did.

### Whether an insert cares that a shape is selected
**Priority: low.** Feasibility: unknown until the sheet arrives.
office-js#2775 (a text-box add deletes the selected shape) and #3698 (a picture
will not insert while one is selected) are both about SHAPES, and this add-in
inserts SLIDES — so the documented repro does not apply and nothing here is
known to be wrong. It is unverified rather than safe, and the preview step now
inserts at a moment when the user may well have something selected.

The probe's `insertWhileSelected` block reports what was selected and whether
the insert landed anyway. It never SELECTS anything: the obvious workaround is
`setSelectedShapes`, which is the one call in this family with a measured
history of wedging the host, so nothing is built for this until there is a
reason.

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

### Charts and SmartArt
**Priority: medium.** Feasibility: low to medium.
Text lives in `charts/chart*.xml` and `diagrams/data*.xml` with embedded
workbooks. Until it is built, the pane must say so out loud rather than skipping
a field the user placed.

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
