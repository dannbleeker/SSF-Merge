## What this changes

<!-- What the reader gets, not what the diff does. -->

## Why

<!-- The problem, and the evidence it exists. Link the issue or the round that
     showed it. If a host behaviour decided the design, name it. -->

## Lockstep

Any feature change updates these in the same PR. Tick what applies, delete what
does not.

- [ ] `docs/MANUAL.md` — a new format, tag, or user-visible behaviour
- [ ] `CHANGELOG.md` — under `## [Unreleased]`, in a user's words
- [ ] `docs/BACKLOG.md` — a shipped item is **removed**, a rejected one gets its reason
- [ ] `README.md` feature table

## Proof

- [ ] `npm test` green locally
- [ ] `npm run typecheck` clean
- [ ] New guards were **proven to fail without their fix** — say which assertion went red and why
