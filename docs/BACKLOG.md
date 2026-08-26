# Backlog

The single curated list of what is open. Items graduate from here into a PR and
are **removed when they ship** — the README's feature table and the manual are
where shipped work is described, so anything still listed here is genuinely not
done.

Priority is what it costs the product to be without it, not how interesting it
is to build.

## Next

### Host layer
**Priority: blocking, and it is next once the probe has been run.** Feasibility: medium; the risk is the host, not the code.
Reading the template out of the open deck (`exportAsBase64Presentation` 1.10,
`slide.exportAsBase64` 1.8, `getFileAsync` as the floor), the one-call insert
with a `targetSlideId`, and positional undo with the clamps. Every rule it must
obey is in `CLAUDE.md`.

### Task pane
**Priority: blocking.** Feasibility: high; the design is settled.
Four steps, SSF visual system, English. Layout and copy are approved and drawn
at true width. Preview writes to the real slide and restores from
`SSF_MERGE_TEMPLATE`.

### Manifests
**Priority: blocking.** Feasibility: high.
XML and unified JSON from one source, validated in CI. Requirement floor
`PowerPointApi 1.4` checked at runtime, never declared — a declared floor makes
the add-in vanish from the ribbon with no diagnostic.

## After the first release

### Image fields
**Priority: high.** Feasibility: medium.
`{{Photo|image}}` filling a rectangle through `ShapeFill.setImage` (1.8), or a
picture cloned into the package with its media part. Blocked on probe question 4:
if the host stretches, the engine has to letterbox before sending.

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
- **Bindings as a way round the id refusals.** Asked and answered on the sibling
  project: the host rejects the batch carrying the binding, with a control arm
  proving it was the binding.
- **Detecting charts whose tags were lost to cut/paste.** The count would be
  swamped by the tag writes this host refuses anyway, and would report "this
  host is unwell" rather than "your paste broke a slide".
