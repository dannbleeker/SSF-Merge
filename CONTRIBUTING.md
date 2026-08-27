# Contributing

## Getting set up

```bash
npm install
npm test
```

Node 22 (`.nvmrc`).

## The checks

| Command | What it holds |
| --- | --- |
| `npm test` | The suite |
| `npm run typecheck` | Types, with `noUncheckedIndexedAccess` on |
| `npm run lint` | Type-aware ESLint, including `no-floating-promises` |
| `npm run format:check` | Prettier, on code only. Prose is not reformatted |
| `npm run coverage` | Coverage floors on `src/core` |
| `npm run test:count` | A floor under the number of tests |

CI runs all of them. `npm run lint -- --fix` and `npm run format` fix most of
what the first two find.

## Rules worth knowing before your first change

**`src/core` imports nothing from Office.js.** A test enforces it. That seam is
what lets the engine be tested without a PowerPoint, and PowerPoint on the web
is documented, in `CLAUDE.md`, to lie about object ids.

**A regression test must be proven to fail without its fix.** Break the source,
re-run, confirm the new test goes red *and that the right assertion went red*,
then restore. A guard that passes against the unfixed code is decoration. Say in
the PR which assertion you saw fail.

**Documentation lands in the same change as the feature.** `test/docs.test.ts`
reads the format kinds and tag keys out of the source and fails when the manual
has not caught up.

**Backlog items are removed when they ship, not ticked.** Anything still listed
is genuinely open.

**Test files are named by topic, never by increment.** No `batch-3.test.ts`.

**Sample data is invented.** The repository is public.

## Host behaviour

Before writing anything that talks to PowerPoint, read the host rules in
`CLAUDE.md`. They are recordings from a sibling project's field rounds, not
opinions, and several of them are the reason the architecture is what it is.
`docs/SIBLING.md` is the fuller ledger: every finding that came across, what was
done about it, and the ones that are **no exposure** here — which is worth
reading before concluding that something has been missed.
