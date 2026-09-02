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
| 2026-08-26 | Dependabot #9: `@types/node` 22.20.1 → 26.2.0 | Not a security alert. The types describe Node 26; this repo pins Node 22 in `.nvmrc`, in `engines`, and in the Pages workflow. Types ahead of the runtime is the silent direction — a call typechecks and is missing at run time. **What was actually measured**: the whole gate passes under `@types/node` 26, and five candidate Node-26 APIs (`assert.partialDeepStrictEqual`, `process.threadCpuUsage`, `process.features.typescript`, `util.getSystemErrorMessage`, `process.ref`) all exist in Node 22.22 too. No break was reproduced, and this entry says so rather than implying one. | **Refused**, and the major pinned in `dependabot.yml` so it stops being re-proposed. Alignment is free and the risk is asymmetric. Raise it deliberately with `.nvmrc` when the repo moves Node. The Pages workflow also stopped hardcoding `node-version: 22` and reads `.nvmrc` like CI, so the version has one home. |
| 2026-08-31 | Dependabot #190: `typescript` 5.9.3 -> 7.0.2 | Not a security alert. TypeScript 7 is the native Go port and **ships no stable programmatic compiler API**; 7.1, targeted autumn 2026, restores it. `typescript-eslint` reads that API, so every published version caps its peer at `>=4.8.4 <6.1.0` — checked against the registry, including `latest` 8.69.0 and the 8.69.1-alpha canary. `npm ci` therefore died with ERESOLVE in 6 seconds, before any type was checked, which is why both CI jobs went red. `typescript-eslint#12518` was closed as *not planned*: the fix is on the TypeScript side. **What was actually measured**: the intermediate 6.0.3 installs clean against the existing toolchain and the whole gate passes on it — typecheck, lint, and the whole suite, which was 1201 tests on 2026-08-31. One real break surfaced and was fixed: TS6 made naming files on the command line while a `tsconfig.json` exists an error (TS5112), which broke `test/probe.test.ts`; the flag that error recommends, `--ignoreConfig`, resolves it. | **Refused for 7.x, taken to 6.0.3.** Microsoft's own path is 5 -> 6 -> 7, because 6 deprecates what 7 removes, so this is the step that had to happen anyway rather than a detour. The 7.x major is ignored in `dependabot.yml` with one un-ignore condition written beside it: a typescript-eslint release whose peer range admits 7.x. #190 closed as superseded. |
| 2026-09-02 | Dependabot #209: the actions group, `checkout` 4→7, `setup-node` 4→7, `upload-artifact` 4→7, in `pane-audit.yml` and `sibling-watch.yml` | Not a security alert. The PR's green checks prove nothing about it: both workflows it changes are `schedule` + `workflow_dispatch` only, so neither runs on a pull request — the same gap this log recorded for the Pages actions on 2026-08-26. **What was actually measured**: both were dispatched by hand on `main` after the merge. `pane audit` **succeeded**. `sibling watch` **failed**, and did NOT fail because of the bump — it was dispatched once on the un-bumped `main` first, and failed there too, so there is a control rather than a guess. | **Taken.** The branch was `BEHIND` and had to be updated before GitHub would merge it, which is why the first merge attempt was refused rather than silently queued. The `sibling watch` failure is pre-existing and is the guard working: *"no keys read from PENDING_QUESTIONS — it was renamed, moved, emptied or reformatted upstream. A sweep that cannot read the table must say so rather than report a quiet week."* The sibling repo moved the table; this repo's watcher is refusing to report a quiet week over it. Logged here, and left for whoever next touches the sibling. |
