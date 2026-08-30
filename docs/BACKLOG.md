# Backlog

The single curated list of what is open. Items graduate from here into a PR and
are **removed when they ship** — the README's feature table and the manual are
where shipped work is described, so anything still listed here is genuinely not
done.

Priority is what it costs the product to be without it, not how interesting it
is to build.

**One list this does not hold.** What an AppSource listing would need is in
[PUBLISHING.md](PUBLISHING.md). Those are not product features — a privacy
policy, a seller account, a screenshot at the store's size — and mixing them in
would put "open a Partner Center account" beside "read a semicolon paste" as
though they were the same kind of thing. This file is the curated list of what
the ADD-IN is missing; that one is the curated list of what a SUBMISSION is
missing. Neither repeats the other.

## Next

**Nothing is waiting on an answer any more.** This section held the probe
questions, and every one it was built to ask has been put and answered — the
last of them, the aspect-ratio one, on 2026-08-28. `docs/PROBE.md` has what they
said.

The first release is out, so the section below is not "later" any longer: it is
the work, in priority order.

## After the first release

**Empty.** The last item here — an Excel table through Microsoft Graph — was
dropped by the owner on 2026-08-30 and is recorded below. Nothing is queued
behind it.

An empty backlog is a statement, not an oversight: this add-in does one thing,
and the two most recent entries were both dropped because they would have made
it do a second. The next item comes from somebody reporting a wall they hit,
which is what the rejected entries below are each waiting on.

## Rejected — do not re-propose

- **Excel via Microsoft Graph.** Dropped by the owner on 2026-08-30: the
  security cost is too large for what this tool is.

  A named table on OneDrive or SharePoint, read through
  `/workbook/tables/{id}/rows`, was the highest-priority item on this list for
  the life of the backlog, and the research that killed it did not find it hard
  to build. The engine needed no change at all — `RecordSet` is the one shape
  every source answers with, and its own docstring already named a Graph-read
  table as a future one. The manifest needed no change either, so no re-install:
  nested app authentication is brokered by the Office host, and Microsoft's own
  reference sample ships the same `<Permissions>ReadWriteDocument</Permissions>`
  this add-in already has, with no `WebApplicationInfo` and no requirement-set
  declaration. The measurable costs were a dynamically-imported
  `@azure/msal-browser` (+72.6 KB gzipped against a 74.0 KB pane), roughly 760
  lines across sign-in, reader and pane, and an Azure app registration.

  What settled it is the permission, and it is not a detail that better wording
  could soften. **Microsoft publishes no read-only path to a workbook's rows.**
  Every Excel endpoint a reader would call — list rows, get range, list tables,
  list worksheets, get table — publishes delegated `Files.ReadWrite` as its
  *least privileged* permission with "Not available" in the higher column. So a
  merge that only ever reads has to ask the user to let it write, and the
  consent dialog says so in Microsoft's words, not ours: *"Have full access to
  all files user can access."*

  That is the opposite of what `SECURITY.md` is able to promise today — no
  network calls, no backend, no server to breach, no log to leak, and a manifest
  that asks for nothing beyond its own host. A tool whose data is usually names,
  addresses, salaries and patient identifiers trades on exactly that, and
  trading it for a more convenient step 2 is a bad exchange. The paste box is
  one keystroke from a range selected in Excel and needs no token, no tenant and
  no trust.

  Three limits are worth keeping written down, because they narrow who the
  feature would have reached even at full price. Workbooks on **personal**
  OneDrive are refused by the Excel API outright ("only the files stored in
  business platform are supported"), so it was a work-account feature only. On
  Office **on the web** — the only host this project has ever validated against
  — nested app authentication 1.1 is supported "only for documents that are
  opened from Microsoft SharePoint Online and OneDrive", so a deck opened from
  disk would get no silent sign-in there. And `Files.ReadWrite.All` needs no
  admin consent by Microsoft's default, but a tenant that has turned user
  consent off turns the dialog into a ticket for IT.

  Reviving it means Microsoft publishing a read-only permission for the Excel
  API. That is the specific thing to watch for — not user demand, which was
  never the doubt.

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
  request, not the easy half's availability. This entry used to add that
  Excel-via-Graph would supply the authentication for free and change the
  calculation; that item was itself dropped on 2026-08-30, so the token this
  half needs has no other reason to exist and would have to be justified on its
  own.

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
