# Security

## Reporting

Report a vulnerability privately through GitHub's
[security advisories](https://github.com/dannbleeker/SSF-Merge/security/advisories/new)
rather than in a public issue.

## What this add-in touches

Worth stating plainly, because merge data is usually personal: names,
addresses, salaries, patient identifiers.

- **The merge runs on the user's own machine.** The engine is client-side
  TypeScript. Rows pasted or read from a local file are never uploaded.
- **It makes no network calls at all.** There is no `fetch`, no
  `XMLHttpRequest` and no socket anywhere in `src/`. The pane and its assets are
  served from `ssf-merge.struktureretsundfornuft.dk` and nothing else is
  requested after that.
- **The add-in has no backend.** There is no server to breach and no log to
  leak.
- **The manifest asks for nothing beyond its own host.** No `AppDomains`, so the
  pane cannot navigate off it, and no `WebApplicationInfo`, so it requests no
  Microsoft identity or Graph scope.

Two backlog items — an Excel table through Microsoft Graph, and saving one file
per recipient to OneDrive — would each add a network destination and a token.
**Neither is built.** This page said they were, in the present tense, until
2026-08-29. When either lands, this section is part of the work.

## Handling untrusted input

A .pptx and a pasted table both arrive from outside, and a user can be sent
either. What the code does about it:

- **A pasted value reaches XML as text and only as text.** It is assigned to a
  DOM node's `textContent` and escaped by the serialiser on the way out. There
  is no `escape()` for a caller to forget, and nothing builds markup by string
  concatenation.
- **A column name is data, not a property.** Rows are objects with a null
  prototype and every read goes through `hasOwnProperty`, so a column called
  `__proto__` or `constructor` merges like any other.
- **A relationship target is not trusted to name a part for deletion.** Removing
  a slide collects that slide's notes page and comments; a target resolving
  outside `ppt/` is left alone.

Each of those is executable, and each is executed: `test/security.test.ts`. A
claim only a document makes is a claim nobody re-checks.

## The sweep of 2026-08-29

What was looked at, and what it said. Filed from
[the backlog item](docs/BACKLOG.md) asking for exactly this.

| Surface | Finding |
| --- | --- |
| Hostile values through a merge | **Sound.** `<b>bold</b>`, `a & b`, `]]>break`, a comment, a processing instruction and an already-escaped entity all round trip exactly once. |
| Injection into the package | **Sound.** A value closing its own `<a:t>` and opening a new element arrives as text; every part still parses and no injected element appears. |
| `__proto__` as a column name | **Sound.** Own property on a null-prototype row, merges onto the slide, pollutes nothing. |
| Relationship targets → part paths | **Two findings, both fixed below.** |
| Traversal out of the package | **Not possible.** `..` past the root is clamped, and the result is always a package-relative name. Nothing here touches a real filesystem. |
| Pasted text into the pane | **Sound.** `render.ts` writes `textContent` and never `innerHTML`; held by a test in `pane-render.test.ts`. |
| Dependency surface | **Two runtime dependencies**, `@xmldom/xmldom` and `jszip`. Alerts are triaged and recorded in `docs/DEPENDENCY-ALERTS.md`. |

### What it found

Both findings are the same shape: **removing a slide decides what to delete from
relationships that came out of the deck**, and a deck can be sent to somebody.

**One — the slide's own notes and comment relationships.** `resolveTarget`
honours a leading `/` and any number of `..`, which is what the format requires.
Removing a slide fed its answer straight to `removePart`, so a slide carrying a
`notesSlide` target of `/[Content_Types].xml` had the merge delete the one part
a presentation cannot open without. Fixed by refusing to remove a part
resolving outside `ppt/`, where a notes page and a comment part always live.

**Two — a chart's or diagram's own relationships.** Found by looking for
siblings of the first, which is the step that gets skipped. Removing a slide
also sweeps the parts its charts and diagrams own. The top of that list was held
to an allowlist; the **children were not**, so a chart whose relationships named
`/ppt/presentation.xml` had that part counted as something the chart owned.
Nothing else in the package refers to it — its only referrer is the root
`_rels/.rels`, which the referrer scan does not read — so it was swept.

This one is worse than the first. `/ppt/presentation.xml` was deleted
**silently**: the merge finished, reported success, and produced a file that
cannot open. `/[Content_Types].xml` was deleted and then threw. Fixed by holding
the children to the same allowlist as the parent — a chart or diagram owns its
styling, its workbook, its diagram parts and its media, and nothing else.

Severity for both is low: they corrupt the user's own output, with no execution,
no persistence and nothing exfiltrated. Both tests were checked by removing the
guard and watching them go red, so they fail for the reason they claim.

**Left standing, and worth knowing:** the referrer scan reads only paths
containing `/_rels/`, which excludes the root `_rels/.rels`. So a part referenced
*only* from there still looks unreferenced. The allowlist now keeps such a part
out of the sweep's reach, so this is latent rather than exploitable — but it is a
wrong answer waiting for someone to widen the allowlist.

### What this sweep did not cover

- **The host.** Office.js, WebView2 and PowerPoint itself are outside it.
- **Delivery.** How the pane is served and how the manifest reaches a tenant
  were not examined.
- **The two unbuilt network features** above. They have no code to read yet.
- **Denial of service.** A deliberately enormous or deeply nested package was
  not tried; the failure there is a slow or failed merge on the user's own
  machine, which is why it ranked below the rest rather than being dismissed.
