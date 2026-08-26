# Changelog

Notable changes to SSF Merge. Newest first.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project uses [semantic versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- The package layer: reading and writing a .pptx as parts, relationships,
  content types and the slide id list, with base64 in and out.
- Slide cloning, with its own relationships, content type, presentation
  relationship, slide id, notes page and a **fresh `p14:creationId`** per copy.
- Tags written into the file as `ppt/tags/tagN.xml`, referenced from the slide's
  `custDataLst`, merging with tags another tool already wrote.
- Run-aware placeholder replacement that survives PowerPoint splitting
  `{{FirstName}}` across runs, and leaves every run's properties untouched.
- Table parsing with column type detection, and formatting for numbers, dates
  and case. An ambiguous slash date is refused rather than guessed.
- The documentation set: manual, backlog, changelog, and a test that keeps them
  in step with the code.
- Hosting on GitHub Pages at `ssf-merge.struktureretsundfornuft.dk`.
- A manual trigger on the CI workflow, so a commit whose push event was lost can
  still be tested.

- Prettier and type-aware ESLint, both run in CI. `no-floating-promises` is on,
  because the engine is async throughout and a missed `await` there merges the
  wrong thing quietly rather than throwing.
- Coverage floors on `src/core`, and a floor under the number of tests so a
  reorganisation cannot delete cases silently.
- Dependabot, with a log in `docs/DEPENDENCY-ALERTS.md` where every alert gets a
  written reading, "no exposure" included.
- Contributing guide, security policy, pull request template and code owners.

### Changed

- Vitest moved to 4.x, so the repository does not start two majors behind.

### Fixed

- `Pkg.text` and `Pkg.copyPart` read edits that had not yet been written back to
  the zip. Three tests were passing on the version from disk.
- Five unnecessary type assertions in the package and XML layers, found by the
  new linter.
- The no-Office-imports guard matched the words "Office.js" and "PowerPoint.run"
  in the comments explaining why the engine avoids them, and failed on four
  correct files.
