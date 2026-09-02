# Contributing

## Getting set up

```bash
npm install
npm test
```

Node 22 (`.nvmrc`).

`npm run dev` serves the pane over plain HTTP, which is what you want for
looking at it in a browser. Sideloading it into PowerPoint needs HTTPS on
`localhost:3000`, because that is the URL the dev manifest names — so install a
development certificate once and point Vite at it:

```bash
npx office-addin-dev-certs install
npx vite --https.key <key> --https.cert <cert>
```

The first command generates and trusts the pair and prints where it put them;
put those two paths into the second. It is the only reason this repository needs
a tool it does not depend on, which is why it is written down here rather than
in a script.

## The checks

| Command | What it holds |
| --- | --- |
| `npm test` | The suite |
| `npm run typecheck` | Types, with `noUncheckedIndexedAccess` on |
| `npm run lint` | Type-aware ESLint, including `no-floating-promises` |
| `npm run format:check` | Prettier, on code only. Prose is not reformatted |
| `npm run coverage` | Coverage floors on `src/core` |
| `npm run test:count` | A floor under the number of tests |
| `npm run build:lib` | The library build. `tsc --noEmit` cannot see it fail, and the two scripts that import `dist-lib/` are only ever read as text by the suite — so a broken build here goes green everywhere else |
| `npm run build` | The pane bundle — what a user actually loads. Nothing else here compiles it, so a build that fails only in the bundler goes green in every other check |

CI runs all of them, `npm run coverage` standing in for `npm test`. `npm run
lint -- --fix` and `npm run format` fix most of what the check commands find.

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
is genuinely open. The manual may not mark anything *planned* while the backlog
has nothing open — a guard reads both.

**The manual's screenshots are generated, and nothing re-renders them for you.**
They live in `docs/images`, committed, because GitHub renders a document from
the repository and cannot run a script. If you change the pane's layout or its
wording, re-run them:

```bash
npx vite --port 5199 --strictPort &
node scripts/manual-shots.mjs
```

A test checks that every picture the manual links to is committed and that
every committed picture is linked. Nothing checks whether a picture still
*looks* like the pane — a byte comparison would go red on every Chromium
version bump, which is a gate people learn to ignore.

**Test files are named by topic, never by increment.** No `batch-3.test.ts`.

**Sample data is invented.** The repository is public.

## Host behaviour

Before writing anything that talks to PowerPoint, read the host rules in
`CLAUDE.md`. They are recordings from a sibling project's field rounds, not
opinions, and several of them are the reason the architecture is what it is.
`docs/SIBLING.md` is the fuller ledger: every finding that came across, what was
done about it, and the ones that are **no exposure** here — which is worth
reading before concluding that something has been missed.
