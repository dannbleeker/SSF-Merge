# Publishing to Microsoft Marketplace

What stands between this add-in and an AppSource listing, as a punchlist.

Everything here was checked against the repository on 2026-08-30 rather than
listed from the policies in the abstract — the "already met" section exists so
nobody re-does work that is done, and every gap says what was looked at.

**Read the first question before doing any of it.** If the answer is "colleagues
inside one tenant", none of this list applies.

## The route, and the one worth knowing about

This list is the cost of a **public AppSource listing**, and that is the route
being taken: the publisher account exists, it is verified, and it is enrolled in
the Microsoft 365 and Copilot program.

The alternative is worth knowing anyway, because it stays true. **Central
deployment through the Microsoft 365 admin centre** puts an add-in on
colleagues' ribbons inside a single tenant with no marketplace review at all —
no listing assets, no EULA, no validation cycle. If the audience were ever one
organisation rather than the world, most of what is left below stops being
necessary.

This section opened by calling seller verification the long pole and telling
whoever read it to settle the route first. Both are now moot: verification is
done, and it cost nothing to have. What is left is small enough that the
question no longer gates anything.

## Already met

Checked, so it is not re-litigated later.

| Requirement | State |
| --- | --- |
| `Office.js` from Microsoft's CDN | Met — `appsforoffice.microsoft.com/lib/1/hosted/office.js`, which is the exact URL the policy names |
| HTTPS everywhere | Met — GitHub Pages behind the custom domain |
| Latest non-preview manifest schema | Met — add-in only manifest, schema 1.1, `TaskPaneApp` |
| High-resolution icon | Met — 16, 32, 64 and 80px present; `IconUrl` and `HighResolutionIconUrl` both set |
| Support URL in the manifest | Met, with a caveat below |
| Proportionate permissions | Met — `ReadWriteDocument`, not full control |
| No unexpected transmission of customer data | **Met outright.** The product makes no network calls: `fetch`, `XMLHttpRequest`, `sendBeacon` and `WebSocket` appear nowhere in `src/`. Nothing leaves the pane |
| No pop-ups, no external navigation | Met — hence no `AppDomains` needed |
| Accessibility groundwork | Partly met — live region, focus restore, visible focus in both themes. axe-core (a WCAG 2 A/AA rule engine) runs over every pane state at both widths in both themes, weekly and on every release; no third-party audit |

## What the release workflow will not let you do

Added 2026-09-01, after an audit found each of them open. None needs a decision
from you; they are here so a refusal is recognisable rather than mysterious.

- **A version this repo has not heard of.** The dispatch box is free text and
  nothing looked at it, so `9.9.9` would have been tagged and published against
  manifests saying something else. The pre-flight now checks it against
  `package.json` and against the changelog having a `## [x.y.z]` section to
  read. If it refuses, bump one to match the other and move what is under
  `[Unreleased]` into its own heading.
- **A pane that fails the accessibility audit.** It runs on a release now as
  well as weekly. Minutes of work, which is why it is still not a required
  check on every merge — but a release is manual and rare, and accessibility is
  a certification criterion.
- **A deploy that serves nothing.** `pages.yml` now asks the live site for the
  pane and an icon after publishing. The failure it catches is "installs
  perfectly, blank ribbon button, empty pane", which nobody reports as a deploy
  problem; the release workflow already checked the live URLs, so between two
  releases it was invisible.

## Blocking — a submission fails without these

- [x] **A privacy policy, at an HTTPS URL.** Done —
      `https://ssf-merge.struktureretsundfornuft.dk/privacy.html`, shipping from
      `public/privacy.html` and linked from the landing page so it is reachable
      without a store listing to send people through.

      It names SSF Merge specifically rather than the site in general, names the
      controller — **DBP Invest ApS**, publishing as StruktureretSundFornuft —
      and is separate from any terms, which Microsoft requires: a Terms of Use
      does not count as a privacy policy. Missing or invalid links here are one
      of the top five reasons submissions fail.

      The honest answer was short because the product makes no network calls at
      all. The page says which single item is written to local storage, names
      the key and lists its fields, which is what makes "we collect nothing"
      checkable rather than a slogan.

- [x] **An End User Licence Agreement, at an HTTPS URL.** Done. Microsoft's own
      standard EULA, chosen by the owner on 2026-08-30:
      `https://support.office.com/client/61994a3b-2c87-41c4-a88d-a6455efa362d`.
      It is in the manifest's `termsOfUseUrl` and resolves.

      It previously pointed at the repository's MIT `LICENSE`, which governs the
      SOURCE. A licence telling a developer they may fork the repository is not
      one telling a user what they may do with the add-in.

- [ ] **Bump the manifest version on every submitted update.** `VERSION` in
      `scripts/manifest-source.mjs` is a hand-edited constant, moved when the
      manifest changes rather than when the project releases: `1.0.0.0` →
      `1.0.1.0` on 2026-08-30, for the icon and support-link changes above.
      AppSource requires the version to increment with each update, so a
      submission whose number has not moved since the last one is refused.
      Nothing automates the bump, and nothing needs to until there is a listing —
      the item stays open because remembering is the whole of the process.

      **This entry previously called the split between the manifest's version
      and `package.json`'s a defect. It is not, and the claim was wrong.**
      The decoupling is deliberate, reasoned in that file and pinned by
      `test/manifest.test.ts`: Office rejects anything below 1.0 outright and
      wants four parts, npm wants semver, and a sibling project shipped `0.1.0`
      in four manifests for the life of its repo because every one of its own
      tests passed and nobody had asked Microsoft. Tying them together would
      reintroduce exactly that.

- [x] **The store icons are the sizes AppSource asks for.** Fixed here. For a
      TASK PANE add-in the store icon must be 32×32 and its high-resolution
      partner **64×64**; `HighResolutionIconUrl` pointed at `icon-80.png`.

      80 is the RIBBON size, which is a different image in a different element
      a few lines below in the same generated file — so the wrong number looked
      exactly like the right one, and only AppSource validation would have said
      otherwise. `icon-64.png` was already being built and used by the JSON
      manifest; nothing new had to be drawn. Fixed in the generator rather than
      the generated XML, or `npm run manifests` would have put it back.

- [x] **`ProviderName` must match the Partner Center publisher name.** Done, and
      it needed no change: the manifest says `StruktureretSundFornuft` and so
      does the Publisher Name on the account.

- [~] **Listing assets.** Drafted in [listing/LISTING.md](listing/LISTING.md) —
      summary at 90 characters, description at 378 words, both inside
      Microsoft's bands, a 300×300 marketplace icon, and five screenshots at
      exactly 1366×768, each around 110 KB against a 1024 KB cap.

      **The screenshots are done.** They were placeholders captured from
      `test-kit/`'s deck, which exists to test the engine rather than show it
      off: the photograph was the crop fixture, a blue rectangle with four
      yellow dots, the chart's numbers were 18 and 42, and a real person's face
      was in the corner. There is a purpose-built demo deck now in
      `listing/demo/`, and `test/listing.test.ts` fails if the old pair ever
      returns. Retaking them is two commands, described in LISTING.md.

      **Notes for certification are drafted too**, in the same file. Partner
      Center marks the field required and says twice that omitting it is an
      automatic rejection, and its claims about the button, the test rows, the
      linked deck and the absence of network calls are each checked against the
      file they are claims about.

      **The name is decided: `SSF Merge` stays**, until Microsoft says
      otherwise. Their guidance says avoid unfamiliar acronyms and do not rely
      on a brand to say what a thing does, which is both halves of it — while
      separately ruling out the obvious fix, because a name may not contain the
      Microsoft product name. `Slide Mail Merge` was the suggestion that fitted
      and it is not being applied: the guidance is guidance rather than a
      submission rule, and if certification objects it will say which rule it
      thinks is broken and can be answered then. The cost is paid in the
      summary, which now leads with "Mail merge for slides" because the name
      does not carry those words. `listing.test.ts` keeps the two in step:
      while `DisplayName` does not say "mail merge", the summary must.

- [x] **A Partner Center account, enrolled in the Microsoft 365 and Copilot
      program, with seller verification completed.** Done — the legal business
      profile reads **Authorized** for DBP Invest ApS, with *Microsoft 365 and
      Copilot* as an active program. This was called the long pole and it is
      already behind us.

## Likely to fail review, or to cost a round trip

- [ ] **Test on macOS Safari, and in Chrome and Firefox.** The policy requires
      the add-in work in all of them. Every real-host round so far has been on
      Windows: the web rounds in Edge, and the round of 2026-08-31 in desktop
      PowerPoint, whose pane is WebView2 and so Chromium again. Neither Safari
      nor Gecko has ever opened the pane. A reviewer tests on their own
      platform.

- [~] **Test on a touch-only device.** The policy, quoted exactly, is section
      1120.3: *"All features must work on a touch-only device without a physical
      keyboard or mouse."*

      **Physical** is the word that matters, and this item used to miss it. An
      on-screen keyboard is part of a touch device, so typing and pasting are
      not what the rule forbids — depending on hardware nobody has is.

      What was checked, on 2026-09-02, without a device:

      - **No hover, right-click, double-click or drag anywhere in the pane.**
        No `mouseover`, `mouseenter`, `contextmenu` or `dblclick` handler, no
        `draggable`, and no `:hover` rule that shows or hides anything — hover
        is decoration only. Those are the interactions a touch device cannot
        perform at all.
      - **Touch targets are already gated.** `scripts/pane-shots.mjs` measures
        every hit area across every pane state and fails under 24 px on the
        short side.
      - **Step 1 does not even need the typing.** "Use the slides I have
        selected" takes the block from the host's own selection, which is a
        tap. It is offered only where `getSelectedSlides` exists — PowerPointApi
        1.5, above the 1.2 floor — so on an older host the two boxes remain the
        only way in, and those need the on-screen keyboard.

      Left unticked deliberately: nothing here has been run on a digitizer, and
      the machine has none. This is now a reasoned risk rather than an unknown
      one, and the residual risk is the on-screen keyboard in the paste box.

- [x] **A support page.** Done —
      `https://ssf-merge.struktureretsundfornuft.dk/support.html`, shipping from
      `public/` beside the privacy policy and linked from the landing page.

      **This entry used to call the GitHub repository "compliant" and it was
      wrong.** Microsoft's submission FAQ says the support URL must be a public
      page that does not require authentication, and that "you can't use personal
      social media pages or GitHub repositories", nor "links to files hosted
      online". A repository is out, and so is a file inside one. It was a
      blocker, not a matter of taste.

      The privacy and terms links were the same mistake and were fixed with it.
      Privacy pointed at `SECURITY.md` — a policy for reporting vulnerabilities
      rather than a privacy policy — so a reviewer following it would have read
      the wrong document and concluded there wasn't one. All three are now pages
      on our own site or Microsoft's, and a test refuses any of them going back
      to GitHub.

- [x] **Edit Data on desktop PowerPoint.** Not a policy item, and no longer
      open: it was the one product behaviour no round had verified, and the
      round of 2026-08-31 did it in desktop PowerPoint 16.0.20326. Right-click ▸
      Edit Data on each classic chart opened a workbook holding that row's own
      values rather than the placeholders, and closing the workbook reverted
      nothing on the slide. `docs/TEST-KIT.md` records both decks.

- [ ] **The title rule, read on 2026-09-02, and it appears to catch this name.**
      Section 1100.7, quoted exactly:

      > The title may not include your brand or service unless your offer targets
      > a larger organization or enterprise.
      >
      > - Apps and agents for Microsoft 365 and copilot may not include the brand
      >   or service in the title.

      The first sentence carries an enterprise exception. The bullet under it,
      which is the one that applies to an Office add-in, carries none: apps for
      Microsoft 365 **may not** include the brand in the title, full stop.

      `SSF` is the brand — StruktureretSundFornuft, the publisher name on the
      account and the domain the add-in is served from. On a plain reading,
      `SSF Merge` is a title that includes the brand.

      This does not settle it. The name was kept deliberately on 2026-09-02 —
      see [listing/LISTING.md](listing/LISTING.md) — on the reasoning that the
      guidance is guidance and a real objection would name the rule it thinks is
      broken. This IS the rule it would name, and it is a certification policy
      rather than a style guide, which is the distinction that reasoning rested
      on. Worth deciding again before submitting, with `Slide Mail Merge`
      already drafted and its consequences already worked out.

## Not a problem, but know it before a reviewer meets it

**The manifest declares no requirement sets, deliberately.** The floor is
`PowerPointApi 1.2` and it is checked at runtime, because a declared set the host
lacks makes the add-in vanish from the ribbon with no diagnostic at all — the
reasoning is in `manifest-prod.xml` beside the omission. Nothing in the policies
requires declaring them. The consequence is that the add-in implicitly claims to
run everywhere and refuses politely instead, so a reviewer on an old build meets
a message rather than an absence. That is the better failure, but it is worth
knowing it is deliberate.

## What a real-host round should exercise next

Written on 2026-09-01, after a release pass that changed the undo, the preview
and the date reader. It is here rather than in a message because a list of what
to click goes stale the moment somebody has to remember where it was, and
because the answer to "what should I try" is a property of the last release
rather than of the conversation it came up in.

Everything below has a test; none of it has met PowerPoint.

- **Undo a merge on a deck you have edited by hand.** Merge, then delete two of
  the merged slides yourself and add one of your own where they were, then press
  "Remove these slides". It should take back what it can show is the merge's,
  leave yours, say so, and then stop offering — a second press is not on offer
  once it has met a slide it could not claim. The thing to watch for is any
  slide of yours disappearing; two routes to that were closed this round and
  they are the reason this is first on the list.
- **End a preview after moving the deck under it.** Preview a row, then delete
  one preview slide by hand, or add a slide of your own, and press "Remove the
  preview". The pane must always end up somewhere you can go on from: while a
  preview is up the merge step refuses, so a preview that cannot be dismissed
  is a session that cannot merge.
- **A date pattern with words in it, in your own language.**
  `{{Start|date:den d. MMMM yyyy}}`, `{{Start|date:d. maj yyyy}}`,
  `{{Start|date:yyyyMMdd}}` and a month column spelled the way Excel exports it
  (`1 Mrz 2026`, `1 dez 2026`, `1 out 2026`). The words should print as written
  and the tokens should read.
- **A number with a decimal that ends in 5.** `{{Total|number:2}}` over `1.005`
  should print `1,01`, the way the cell reads in Excel.
- **Pictures from two folders.** Attach one folder, then another. Both sets
  should still be attached, and the pane should say it added rather than
  replaced.

## Sources

- [Microsoft 365 app publishing checklist](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist)
- [Marketplace certification policies](https://learn.microsoft.com/en-us/legal/marketplace/certification-policies) — 1100 general, 1120 Office Add-ins
- [Publish your Office Add-in to Microsoft Marketplace](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource)
- [Top five AppSource validation errors](https://devblogs.microsoft.com/microsoft365dev/top-five-appsource-validation-errors-for-office-add-ins-submissions-march-2023/)
- [Marketplace submission FAQ](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/appsource-submission-faq)
