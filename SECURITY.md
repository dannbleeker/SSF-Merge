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

This is now a decision rather than a state the project happens to be in. The
two backlog items that would each have added a network destination and a token
— an Excel table through Microsoft Graph, and saving one file per recipient to
OneDrive — were both dropped by the owner, the second of them on 2026-08-30 and
explicitly on the strength of this page: Microsoft publishes no read-only
permission for the Excel API, so reading a table would have meant asking the
user for write access to their files. `docs/BACKLOG.md` has the full reasoning
and what would make either worth revisiting.

Anything that changes that stands as a change to this section as much as to the
code. This page said both features existed, in the present tense, until
2026-08-29 — which is the failure to watch for here: a security page is read by
people deciding whether to trust the tool, and it is the last page anyone thinks
to update.

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
| Relationship targets → part paths | **Three findings, all fixed below.** |
| Traversal out of the package | **Not possible.** `..` past the root is clamped, and the result is always a package-relative name. Nothing here touches a real filesystem. |
| Pasted text into the pane | **Sound.** `render.ts` writes `textContent` and never `innerHTML`; held by a test in `pane-render.test.ts`. |
| Dependency surface | **Two runtime dependencies**, `@xmldom/xmldom` and `jszip`. Alerts are triaged and recorded in `docs/DEPENDENCY-ALERTS.md`. |

### What it found

The first two are the same shape: **removing a slide decides what to delete from
relationships that came out of the deck**, and a deck can be sent to somebody.
The third is why the second worked.

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
`_rels/.rels`, which the referrer scan did not read at the time — so it was
swept.

This one is worse than the first. `/ppt/presentation.xml` was deleted
**silently**: the merge finished, reported success, and produced a file that
cannot open. `/[Content_Types].xml` was deleted and then threw. Fixed by holding
the children to the same allowlist as the parent — a chart or diagram owns its
styling, its workbook, its diagram parts and its media, and nothing else.

Severity for both is low: they corrupt the user's own output, with no execution,
no persistence and nothing exfiltrated. Both tests were checked by removing the
guard and watching them go red, so they fail for the reason they claim.

**Three — the referrer scan could not see the package's own relationships.** The
reason the second finding worked. That scan tested every `.rels` whose path
contained `/_rels/`, which is all of them **except** `_rels/.rels` — the root has
no directory in front of it. So the only referrer of `ppt/presentation.xml` and
of docProps was invisible, and a part named just from there looked unreferenced.

Underneath it sat a second cause: `resolveTarget` mishandled a part at the
package root, the same `lastIndexOf` trap `relsPathFor` already documented. An
empty base dropped a character and prefixed every answer with a slash, so the
root's relationships could not have been read even if the scan had looked.

Both fixed, and the scan now reads the root. Not a vulnerability once the
allowlist is in place — it is a wrong answer that the allowlist was hiding, and
the kind that comes back the day somebody widens one.

### What this sweep did not cover

- **The host.** Office.js, WebView2 and PowerPoint itself are outside it.
- **Delivery.** How the pane is served and how the manifest reaches a tenant
  were not examined.
- **The two unbuilt network features** above. They have no code to read yet.
- **Denial of service.** A deliberately enormous or deeply nested package was
  not tried; the failure there is a slow or failed merge on the user's own
  machine, which is why it ranked below the rest rather than being dismissed.
  **Picked up by the sweep of 2026-08-30 below, which found one** — and the
  entry ranked it right: a frozen tab on the user's own machine.

## The sweep of 2026-08-30

The second pass, aimed at what the first one listed as not covered — denial of
service — and at the parser surfaces it took on trust.

| Surface | Finding |
| --- | --- |
| XML external entities (XXE) | **Not possible, measured.** `@xmldom/xmldom` does not resolve them: `<!ENTITY x SYSTEM "file:///etc/passwd">` comes through as the literal text `&x;` with a parse error logged, and no file is read. |
| Entity expansion (billion laughs) | **Not possible, measured.** Custom internal entities are not expanded at all — a four-level bomb answers `entity not found` and yields three characters. |
| HTML injection in the pane | **Not possible.** No `innerHTML`, `outerHTML`, `insertAdjacentHTML`, `document.write`, `eval` or `new Function` anywhere in `src/`. |
| Network egress | **None.** No `fetch`, `XMLHttpRequest`, `WebSocket` or `sendBeacon` in `src/`, which is what the top of this page claims. |
| User data into an XML **attribute** | **Does not happen.** Every `setAttribute` value is engine-generated — a content type, a part path, a relationship id, a number. No cell value or column name reaches one. |
| A picture's file extension | **Not from the file name.** It comes from a fixed table keyed by the kind `readImage` decides from the file's MAGIC BYTES, so a file called `x".png` cannot reach `[Content_Types].xml`. |
| Dependencies | `npm audit` clean, with and without dev dependencies. |
| Prototype reach through a data cell | **One finding, fixed below.** |
| Backtracking on hostile template text | **One finding, measured and NOT fixed — see below.** |

### One — a date cell could reach `Object.prototype`

`monthFromName` looked a month word up in `MONTH_NAMES`, a frozen object,
keyed by the cell's own text. `Object.freeze` does not remove the prototype
chain, and `dateShape`'s `NAMED_DATE` accepts **any** word of three letters or
more — so `1 constructor 2026` passes the gate, and the lookup answers the
`Object` function. `__proto__` answers `Object.prototype` the same way.

**The outcome was already correct, and only by luck.** A function reaches
`Date.UTC` as a month, the arithmetic is NaN, the date is invalid, and the rule
that an unreadable cell is printed as it stands catches it. Guarded with
`hasOwnProperty` now, because this repo's rule is to guard any table keyed by a
config or data string and the luck was one refactor deep.

The tests beside it are **characterisation tests, not a proof**: with the guard
reverted they still pass, which was checked rather than assumed. They pin the
outcome for the day the luck runs out.

### Two — quadratic backtracking in the placeholder pattern, NOT fixed

`FIELD` in `src/core/merge/text.ts` backtracks super-linearly on template text
that opens a placeholder and never closes it. Measured, on a single paragraph:

| input | time |
| --- | --- |
| `{{` + 5,000 letters | 74 ms |
| `{{` + 10,000 letters | 292 ms |
| `{{` + 20,000 letters | 1,128 ms |
| `{{` + 40,000 letters | 4,447 ms |
| `{{a\|` + 2,000 spaces | 4,232 ms |

Four times the work for twice the input: quadratic. Ordinary slide text is
unaffected — 116,000 characters carrying 6,000 real placeholders match in 3 ms —
and so is a paragraph with many unclosed `{{`, because the character class stops
at the next brace. It needs one long unbroken run inside one paragraph.

**Reachable** from a deck somebody was sent: a `<a:t>` run of forty thousand
characters is legal, and the pattern runs over every paragraph of the template
during the step-1 read as well as during the merge. **The impact is a frozen
tab** — the match is synchronous, so the pane's own call timeouts cannot save
it. No data loss, no escalation, nothing that reaches another user.

**Left unfixed deliberately.** Three candidate patterns were written and
measured against a 30,000-string corpus for exact equivalence, and each one
fixed one shape and made another worse: a lookahead that removes the letters
case takes `{{` + 40,000 spaces from 28 ms to 3,447 ms, and bounding the trim
that fixes the pipe case leaves it. `fieldPattern()` is a public export
returning a `RegExp`, so the linear answer — scan for `{{`, scan for `}}`, split
on the first `|` — is a change to the library's contract and to six call sites,
which is not a thing to do inside a sweep.

What it needs is a decision on that contract. Until then it is written down
here rather than half-fixed, which is the state the three candidates were in.

### What this sweep did not cover

- **A decompression bomb in an embedded workbook.** A chart's `.xlsx` is
  inflated by `jszip` with no size guard. The deck reaching that code was
  exported by PowerPoint from a presentation the user already had open, so the
  bytes are host-produced; the residual case is a hostile chart in a deck
  somebody was sent, and the impact is the same frozen tab as above.
- **The host, delivery, and the two unbuilt network features**, all as before.
