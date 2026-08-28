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

### A chart's numbers, per recipient
**Priority: medium.** Feasibility: medium.
Chart and SmartArt TEXT merges (shipped 2026-08-28). The numbers do not: a
bar's height comes from a `<c:numCache>` cell that has to parse as a number, so
`{{Revenue}}` cannot be written there — a placeholder in one is deliberately
left alone rather than replaced with something the chart cannot plot.

Doing it means a syntax that is not a placeholder in the text (there is nowhere
in the values to put one), writing both the cache and the embedded workbook's
cells, and deciding what a non-numeric cell does. The workbook writer already
exists — the text merge opens and rewrites it — so the missing half is the
syntax and the numeric path, not the plumbing.

### A security sweep, written down
**Priority: medium.** Feasibility: high.
The add-in runs inside PowerPoint, opens whatever deck the user has open, and
takes a table they pasted from somewhere else. None of that is exotic, and a
first look says the surfaces are narrow — which is the argument for doing this
while the answer is still short rather than after something has gone wrong.

Most of the work is confirming and recording a posture, not building one. What
to sweep, and what a first reading already says:

- **Pasted text reaching XML.** User values enter the package through two
  `textContent =` assignments, and `@xmldom`'s serialiser escapes on the way
  out; nothing builds XML by string concatenation today. Confirm that holds, and
  that a value containing `<`, `&`, `]]>` or a lone surrogate makes a round trip
  without corrupting the package — the harm here is a deck that will not open,
  which is the same harm as a bad merge.
- **The input deck is untrusted.** Zip entry names become part paths. Confirm a
  crafted name — `../`, absolute, absurdly long — cannot write outside the
  package, and state the parser's posture on entity expansion for what it reads.
- **Pasted text reaching the pane.** `render.ts` uses `textContent`, never
  `innerHTML`, and one test holds it. Confirm that guard covers every path
  showing a column name, a file name or a placeholder, rather than only the one
  it was written against.
- **The dependency surface.** Two runtime dependencies, `@xmldom/xmldom` and
  `jszip`. Decide how they are watched and write it down; `DEPENDENCY-ALERTS.md`
  is the obvious home.
- **What the manifest permits.** There is no `AppDomains` element, so the pane
  cannot navigate outside its own host. That is the right answer, and it should
  be the deliberate one rather than the default one.

The output is a page saying what was checked and what the answers were, so the
next person inherits a posture instead of repeating the reading.

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
