# Backlog

The single curated list of what is open. Items graduate from here into a PR and
are **removed when they ship** — the README's feature table and the manual are
where shipped work is described, so anything still listed here is genuinely not
done.

Priority is what it costs the product to be without it, not how interesting it
is to build.

## Next

### The preview step — one row on the real slide
**Priority: high.** Feasibility: medium.
Step 3 of the pane is the one thing in the wizard that is not built. It has to
write one record's values onto the real template slide, store what was there in
a `SSF_MERGE_TEMPLATE` shape tag, and put it back — which means editing shapes
through Office.js rather than in the package, the one place this project does
that. The screen says so out loud today and its button carries the user past it;
the merge does not need it.

The other half is picking the block by CLICKING slides rather than typing two
numbers. `getSelectedSlides` is the obvious route and this host is documented to
wedge its whole selection subsystem after `setSelectedShapes`, so it needs a
probe question before it needs code.

### Manifests
**Priority: blocking.** Feasibility: high.
XML and unified JSON from one source, validated in CI. Requirement floor
**`PowerPointApi 1.3`** checked at runtime, never declared — a declared floor
makes the add-in vanish from the ribbon with no diagnostic. This said 1.4 until
the floor was read off the calls the add-in actually makes: everything but
`slide.tags` is 1.2, and `getFileAsync` is a Common API that PowerPointApi does
not gate. `API_FLOOR` in `src/host/capability.ts` carries the table, and a test
fails if it moves.

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
