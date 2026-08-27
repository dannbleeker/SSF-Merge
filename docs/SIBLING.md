# What we know from PowerChart, and how it gets here

Most of what this repo knows about the PowerPoint host was not learned here. It
was learned by [PowerChart](https://github.com/dannbleeker/powerchart), a
sibling add-in by the same author, over 174 real-host rounds against
PowerPoint on the web (counted 2026-08-27). SSF Merge has run zero real merges. Every host rule in
`CLAUDE.md` is a recording from that project.

That is a large and continuing debt, and it does not settle itself. This file is
how it gets paid: the single place PowerChart-derived knowledge lives, what was
done about each item, and the rules that stop a borrowed fact quietly going bad.

## Why this file exists — the drift, measured

On 2026-08-27 there were **44 citations of PowerChart across 24 source, script
and doc files**, every one hand-copied with no link back. Sorting them found two
kinds, and only one of them is a problem:

- **A single run's observation is durable.** "A by-id clean-up reported 45
  successful deletes having removed nothing" happened once, in one round, and
  will be true forever. Most citations are this. Leave them alone.
- **A live counter is a snapshot.** Four places said "174 consecutive archived
  rounds" and one said "passed 174 of 174". PowerChart's archive held exactly
  174 rounds that morning, so all five were right — and round 175 makes every
  one of them wrong, with nothing anywhere to say so.

One had already drifted. `docs/BACKLOG.md` read "of 327 answered batches the
slowest took 31s". The 327 comes from the measurement that reported **29.2s**;
the 31.1s figure comes from a later round with a larger population. A count from
one measurement and a maximum from another, in one sentence, both correct
separately.

## The rule

**A number copied from a live counter carries the date it was taken.**

That single change turns a claim into a recording. "174 of 174" rots the moment
another round runs. "174 of 174, measured 2026-08-27" never becomes false — it
only becomes older, and a reader can judge that for themselves. This is what
every durable citation here already does implicitly by describing one run.

Two corollaries:

- **Where the number adds nothing, drop it.** "Every rung answered in every
  archived round" is stronger than "in all 174" and cannot decay at all.
- **Never pair figures from two measurements.** They read as one finding and
  are not. This is what the batch-timeout sentence did.

PowerChart applies the same rule to itself and states the cost of breaking it:
its own memory file spent several rounds claiming "five sheets" while the table
it referred to held eight.

## How a fact gets here

1. Something is learned in PowerChart — a round, an upstream issue, a reverted
   fix.
2. It is triaged against **this** add-in's surface. The bar is PowerChart's own,
   and it is what keeps a report readable: not "this is interesting about
   Office.js" but **"if this were true, code in this repo would be wrong."**
3. It gets a row below — **including when the answer is "no exposure"**, which
   records that somebody checked. An untriaged item is indistinguishable from an
   unnoticed one.

Anything without a row is unanswered, and comes back.

### The precondition nobody should have to rediscover

Step 2 needs **both repositories checked out in the same session**. A session
holding only SSF Merge cannot read PowerChart's tables, cannot verify a
citation, and cannot do any of this by hand. It can read a sweep's output and
nothing else. Say so when starting work that depends on it.

### Where to read, in order of signal

PowerChart's `CLAUDE.md` is the fullest record and the worst feed — it is over a
thousand lines of prose and diffs into noise. The curated, machine-readable
surfaces are the ones worth watching:

| Source | What it holds |
| --- | --- |
| `scripts/office-js-watch.mjs` → `KNOWN_ISSUES` | Every office-js issue PowerChart has triaged, with what it did about each. We call a subset of the same APIs. |
| `scripts/host-baseline.mjs` → `FAKE_BASELINE`, `KNOWN_DIVERGENCES`, `UNSTABLE_ANSWERS`, `PENDING_QUESTIONS` | What that project knows about this host, already keyed and worded for a diff. |
| `docs/BACKLOG.md` rejected list | Decisions not to do things, with reasons. Cheaper to read than to re-derive. |
| `rounds/*.json` | One file per real-host round. Most say nothing new; the pooling tools are what read them. |

## The ledger

Every row is a PowerChart finding and what SSF Merge did about it. `no exposure`
is a real answer and the most common one worth writing down.

| PowerChart finding | Source | What we did |
| --- | --- | --- |
| A slide the run just added does not resolve by `slides.getItem(id)`; a by-id clean-up reported 45 deletes and removed nothing | field rounds; `deleteSlideByPosition` | **Adopted as architecture.** Merge metadata goes into the package before the insert; undo is positional with clamps (`src/host/undo.ts`). |
| Tag writes through a shape proxy are refused — 46 `InvalidParam passed to GetItem(id)` in one run | field round | **Adopted as architecture.** Tags are written into `ppt/tags/tagN.xml` before the insert, so nothing is asked (`src/core/pptx/tags.ts`). |
| `insertSlidesFromBase64` without `targetSlideId` inserts at the FRONT — 37 slides ahead of a title slide | field round | **Guarded.** `insertDeck` reads `lastSlideId()` first (`src/office/powerpoint.ts`). |
| A queued call that raises nothing has not necessarily happened | field rounds | **Adopted.** Every write is verified by a deck delta, never by the absence of an error. |
| A stall is death, not slowness: no abandoned call ever answered late | 17 abandoned calls over 9 rounds | **Adapted, and it points the other way here.** Our probe's third sheet timed out on an insert whose delta showed both slides had landed, so budgets differ per call and the DELTA decides. See `src/host/timeout.ts`. |
| Collection loads over ~50 items answer short (office-js#4272); PowerChart pages every collection read at 20 | `ID_PAGE` in its renderer | **Adopted.** `ID_PAGE = 20` in `src/office/powerpoint.ts`, same number for the same reason. Asked directly by the `deckRead` probe question. |
| A `SlideRange` id lacks the deck's `#suffix` (office-js#2474) | its selection code | **Adopted.** `deckIdForSelectedSlide` matches by prefix and refuses two matches (`src/host/capability.ts`). |
| `setSelectedShapes` wedges the web host's selection subsystem (office-js#3083, #3698) | its selection ladder | **Avoided entirely.** This add-in never calls it. `getSelectedSlides` is read-only and measured safe. |
| Setting shape text through Office.js re-authors it (office-js#5858); a shape-by-shape run took 680 seconds and shipped duplicate slides | its own architecture note | **Adopted as architecture.** The whole package-merge design is the answer to this. |
| A contract gate read "never asked" as an answer | its `no-scratch-slide` handling | **Guarded before we could repeat it.** `tagVerdict` returns `unknown` with a detail saying the read fell on a slide the probe did not write (`src/host/verdicts.ts`). |
| `import.meta.url.endsWith(...)` never splits a backslashed path — three tool CLIs silently exited 0 on Windows for months | its `is-main` fix | **Adopted.** `scripts/is-main.mjs` compares file URLs. |
| A manifest below `<Version>1.0` is rejected by Office; four manifests shipped `0.1.0` for the life of that repo | its manifest job | **Guarded offline.** `scripts/manifest-rules.mjs` fails below 1.0; version is `1.0.0.0`, independent of the npm package. |
| Every gate reads the working tree; a user downloads the RELEASE, and the two diverged for twelve days | its v0.1.0 release | **Guarded.** `scripts/release-assets.mjs` checks the attached assets against the docs. |
| `npm audit`'s proposed remedy moved a runtime dependency back three major versions | its dependency log | **Adopted as a rule.** `docs/DEPENDENCY-ALERTS.md` step 3. |
| One answer sheet is evidence about the host *in that minute*; the same build answers differently minutes apart | its `UNSTABLE_ANSWERS` | **Adopted as doctrine.** `docs/PROBE.md` says to run the probe twice when an answer decides something expensive. |
| A trace line named for an outcome it is written before knowing is false on exactly the runs that matter | its `batch committed` → `batch issued` | **Pending.** Constrains the trace in the debug plan; nothing to change until that ships. |
| A value recorded only on failures cannot discriminate | its `idleMs`, `afterAnswering` | **Pending.** Same. `withTimeout` will record on both populations. |
| Two time series in one artefact must share one clock origin | its probe stamping 7.9s off the trace | **Pending.** Same. |
| Duplicate `<p:cNvPr id>` within a slide triggers PowerPoint's repair dialog and is invisible to schema validation | its `verify-deck.mjs` | **No exposure today, and worth stating why.** We clone whole slide parts, so ids stay unique within a slide. It becomes exposure the moment anything edits a shape tree, and it is on the list for the package self-check. |
| `getImageAsBase64` on a freshly added slide has killed the tab five distinct ways | its rasterise rules | **No exposure.** Nothing here rasterises anything. |
| Two `slides.add()` calls in quick succession kill PowerPoint on the web | field round | **No exposure.** This add-in never calls `slides.add`; slides arrive through one `insertSlidesFromBase64`. |
| Drawing cost grows quadratically with the shapes already on a slide | its degradation experiment | **No exposure.** There is no per-shape draw loop; a merge is one insert. |
| Grouping, tag-settle and id-refusal recovery | most of its renderer | **No exposure.** No shape-level work here at all. |
| Waiting after `slides.add()` (office-js#2903) cost 18 of 19 probe answers in one round | tried and reverted there | **Rejected before trying.** `docs/BACKLOG.md` rejected list. |
| Bindings as a route around id refusals — the host rejects the batch carrying the binding | its probe, with a control arm | **Rejected before trying.** `docs/BACKLOG.md` rejected list. |
| A slide add whose sync never resolves though the slide lands (office-js#1650) | its bounded slide-adds | **Adopted as doctrine.** We never call `slides.add`, but `insertSlidesFromBase64` is the same shape and gets the same answer: the deck delta decides. |
| `addTextBox` deletes the SELECTED shape on the web (office-js#2775) | its `dropShapeSelection` | **No exposure to the call, but the class is ours.** The preview inserts when the user may have something selected, so `insertWhileSelectedProbe` asks rather than assuming. |
| The web uppercases tag KEYS internally and needs the uppercased spelling to read them back (office-js#6079) | its tag writer | **Relevant, and already safe by luck.** Every key we write is uppercase (`SSF_MERGE_RUN`, `SSF_MERGE_RECORD`, `SSF_MERGE_BLOCK`). A lowercase key would go into the package fine and be unreadable on the web. |
| `Slide.exportAsBase64` omits modern comments and `ppt/authors.xml` (office-js#6867) | its round evidence | **A finding for us that the sibling correctly marked no exposure — see below.** |
| Shape tags are lost on cut/paste on the web (office-js#3784) | its triage | **Documented, not guarded.** A merged slide cut and pasted into another deck loses its run tag, so undo cannot find it. Caveat in `docs/MANUAL.md`; detection is refused for the sibling's reason. |
| Inserted content may appear in the slide PREVIEW but not the main view (office-js#6498) | its visibility gate | **Relevant as a support answer.** A user reporting missing merged slides may be seeing this, and the deck delta will say they landed. Nothing here can read the canvas to tell them apart. |
| `PowerPoint.run` batching fails to load properties reliably after `sync()` (office-js#6363) | its central failure | **Highly relevant.** `deckSlideIds` batches `load("id")` across twenty `getItemAt` handles and reads them after one sync — precisely this shape. Asked by our `deckRead` probe's `empty` arm. |
| `getcount-populates-same-sync` — the count is right while the list is empty | its `FAKE_BASELINE` | **Relevant.** Why the paging loop trusts the scalar `getCount()` and never a collection load. |
| `getitemat-past-end` — what the host does past the end of the collection | its `FAKE_BASELINE` | **Relevant.** Bounds both `deckSlideIds` (paging by index) and `undoInsert` (deleting by index). |
| `which-end-a-short-read-drops` | its `FAKE_BASELINE` | **Relevant.** A short read that is not a prefix makes a slide NUMBER wrong, not merely a list shorter. Our probe asks the same thing as `prefixOk`. |
| `how-many-collection-reads-a-context-survives` | its `PENDING_QUESTIONS` | **Relevant.** `deckSlideIds` takes one `PowerPoint.run` per page deliberately, so no context accumulates reads. Recorded so the reason survives a refactor. |
| `delete-then-lookup` — whether a deleted slide still resolves | its `FAKE_BASELINE` | **Relevant as doctrine.** Exactly what made by-id clean-up unsafe. Undo deletes by position, highest index first, and re-counts rather than believing the call. |
| `scratch-slides-returned` — whether a probe gets its slides back | its positional sweep | **Adopted.** Our probe's sweep is positional and triple-clamped, each clamp proven load-bearing in `test/undo.test.ts`. |

### The one open risk this sweep surfaced

**`exportAsBase64Presentation` may be dropping parts, and nobody has checked.**
office-js#6867 reports that `Slide.exportAsBase64` omits modern comments and
`ppt/authors.xml`. PowerChart marked it no exposure and was right to: it calls
that API to get a PICTURE of a slide. We call the presentation-level export to
read the TEMPLATE WE THEN CLONE, so any part the export drops is a part every
merged slide is missing — silently, in a file that opens cleanly.

Different call in the same family, and the presentation-level one has never been
tested for it. It wants a probe question before the first real merge on a deck
that has comments. Recorded here rather than fixed because guessing at which
parts are affected would be worse than measuring.

### What reading it found in OUR code

The archaeology was aimed at instruments and turned up a product defect. **On
every host below PowerPointApi 1.10 the merge would have inserted the user's
entire presentation a second time** — the `file` route returns the whole deck
and only the template block was being removed from it. Fixed, with both routes
now held by `test/office-merge.test.ts`.

Worth recording as a transfer even though no PowerChart finding names it,
because the rule that caught it is one of theirs: **measure the artefact you
hand over, not your intent.** The count sent to the host was `result.slides.length`
— what the plan believed it built — where the package's own `sldIdLst` was the
only honest answer, and on that route the two disagreed by the size of the
user's deck.

## What we learned that PowerChart has not

PowerChart is read-only by the owner's instruction, so findings that run the
other way wait here. These are ours, from this project's three answer sheets and
its own design work.

- **A malformed fixture answers `InvalidArgument`, not "the scheme is wrong".**
  Our first sheet reported that the metadata scheme needed rethinking, on the
  strength of a read that had landed on the user's own title slide because the
  insert in front of it threw. The fix was two control arms, not reasoning. The
  general form — *an arm that proves the question was asked* — is a probe
  design rule PowerChart's sheets would also benefit from.
- **On the insert path the DELTA points the opposite way to "a stall is
  death".** An insert that timed out at thirty seconds had landed both slides.
  Reading the raise as decisive produced three false statements in one run.
- **The package-insert surface is almost unmeasured.** `getFileAsync` slicing on
  a large deck, `exportAsBase64Presentation`'s ceiling, and what
  `insertSlidesFromBase64` does with a 400-slide package are questions
  PowerChart has never had to ask.

## Keeping this file honest

It is subject to its own rule. The 44-citation count and the 174-round figure
above are dated because they are counters, and they will be stale rather than
wrong. Re-count before quoting either.

**The sweep runs weekly** (`.github/workflows/sibling-watch.yml`, Mondays, an
hour after PowerChart's own office-js sweep so a finding it triages that morning
is in its tables before this looks). `scripts/sibling-watch.mjs` reads the
curated tables above and reports anything with no row in its `TRIAGED` map —
one issue, reopened and updated, never one per sweep. Run it yourself with
`npm run sibling-watch`.

Raw file reads only, never the GitHub API: both repositories are public, so it
needs no token and runs in CI and in an agent session alike. It never imports
the sibling's code — a weekly job that executes a file fetched over the network
is a supply chain, not a sweep.

Three exit codes, and the third is the one that matters: 0 when everything has a
row, 3 when something does not, and anything else when the sweep itself broke. A
table renamed upstream throws by name rather than matching nothing, because
"nothing new, every Monday, forever" is indistinguishable from a quiet week and
is exactly the failure this file exists to prevent.

**`TRIAGED` is the source of truth and this prose is downstream of it.** Every
reason opens with a verdict from a closed vocabulary — `NO EXPOSURE`, `ADOPTED`,
`RELEVANT` — and `test/sibling.test.ts` fails when a finding we ACTED on has no
line here. It found thirteen missing rows the first time it ran, against a
ledger written the same morning.
