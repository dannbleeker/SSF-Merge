# Dependency alerts

Every alert gets read and the reading gets written down here, **including "no
exposure"** — that entry is what records that somebody looked.

A standing red count with nobody looking is worse than no scanner at all: it
trains everyone to scroll past the one that matters. If this file is empty and
the alert count is not zero, the process has stopped working.

## How to triage one

1. Is the vulnerable code path reachable from this package? Most alerts on a
   build-time dependency are not.
2. Is there a fix that does not move a major version?
3. **Never take `npm audit`'s advice unread.** On the sibling project its
   proposed remedy for one alert was to move a runtime dependency back three
   major versions.
4. Record the decision below with the date, whatever it was.

## Log

| Date | Alert | Reading | Action |
| --- | --- | --- | --- |
| 2026-08-26 | none open at the time Dependabot was switched on | Baseline entry. Two runtime dependencies, `jszip` and `@xmldom/xmldom`, both read package bytes the user supplies, so both are genuinely reachable and worth watching. | Watching |
| 2026-08-26 | Five GitHub Actions majors: `checkout` 4→7, `setup-node` 4→7, `configure-pages` 5→6, `upload-pages-artifact` 3→5, `deploy-pages` 4→5 | Not security alerts, just Dependabot's first sweep finding the actions pinned by major. They arrived as five separate pull requests because the actions ecosystem had no group; that is fixed, and the next sweep will open one. | **Taken**, all five together, rather than waiting a month for the grouped sweep. CI proves `checkout` and `setup-node`; the three Pages actions are not exercised by CI at all, so the Pages workflow was dispatched by hand afterwards to prove the deploy still works. Dependabot's five pull requests closed as superseded. |
