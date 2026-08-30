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
| Accessibility groundwork | Partly met — live region, focus restore, visible focus in both themes. Not audited against a standard |

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

- [ ] **An End User Licence Agreement, at an HTTPS URL.** Also required.
      Microsoft publishes a standard EULA that may be used if your own counsel
      agrees. The repository is MIT-licensed, which is not the same thing and
      does not satisfy this.

- [ ] **One version number.** `manifest-prod.xml` says `1.0.0.0` and
      `package.json` says `0.2.3`. The policy requires the version to increment
      on every update, and today there is no single source of truth for what it
      is. This is worth fixing whichever route is taken.

- [x] **`ProviderName` must match the Partner Center publisher name.** Done, and
      it needed no change: the manifest says `StruktureretSundFornuft` and so
      does the Publisher Name on the account.

- [ ] **Listing assets.** At least one screenshot is required, plus a short
      description and a long one with HTML formatting. The pictures in
      `docs/images` are 380px pane captures — right for the manual, too small
      for a store listing.

- [x] **A Partner Center account, enrolled in the Microsoft 365 and Copilot
      program, with seller verification completed.** Done — the legal business
      profile reads **Authorized** for DBP Invest ApS, with *Microsoft 365 and
      Copilot* as an active program. This was called the long pole and it is
      already behind us.

## Likely to fail review, or to cost a round trip

- [ ] **Test on macOS Safari, and in Chrome and Firefox.** The policy requires
      the add-in work in all of them. All three real-host rounds so far were
      Edge, on Windows, on the web. A reviewer tests on their own platform.

- [ ] **Test on a touch-only device.** "All features must work on touch-only
      devices without keyboard/mouse." This flow needs two typed slide numbers,
      a pasted table and a file picker, and none of that has been tried without
      a keyboard.

- [ ] **Decide whether a GitHub repository is the support page.** It is
      compliant — an HTTPS URL, and not an email address, which is what the
      policy forbids — but a README is a thin answer to "where do I go when this
      breaks".

- [ ] **Edit Data on desktop PowerPoint.** Not a policy item. It is the one
      product behaviour no round has ever verified, and desktop is exactly where
      a reviewer will be. See `docs/TEST-KIT.md`.

- [ ] **Read the title rule.** Titles may not include a brand or service name,
      with an exception for add-ins aimed at large organisations. Whether "SSF
      Merge" is caught by that is worth asking before submitting rather than
      after.

## Not a problem, but know it before a reviewer meets it

**The manifest declares no requirement sets, deliberately.** The floor is
`PowerPointApi 1.2` and it is checked at runtime, because a declared set the host
lacks makes the add-in vanish from the ribbon with no diagnostic at all — the
reasoning is in `manifest-prod.xml` beside the omission. Nothing in the policies
requires declaring them. The consequence is that the add-in implicitly claims to
run everywhere and refuses politely instead, so a reviewer on an old build meets
a message rather than an absence. That is the better failure, but it is worth
knowing it is deliberate.

## Sources

- [Microsoft 365 app publishing checklist](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist)
- [Marketplace certification policies](https://learn.microsoft.com/en-us/legal/marketplace/certification-policies) — 1100 general, 1120 Office Add-ins
- [Publish your Office Add-in to Microsoft Marketplace](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource)
- [Top five AppSource validation errors](https://devblogs.microsoft.com/microsoft365dev/top-five-appsource-validation-errors-for-office-add-ins-submissions-march-2023/)
- [Marketplace submission FAQ](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/appsource-submission-faq)
