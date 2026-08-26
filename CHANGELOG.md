# Changelog

Notable changes to SSF Merge. Newest first.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project uses [semantic versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed

- Questions three and four could not be reached: a shape proxy does not survive
  a `context.sync()` on PowerPoint for the web, which answers `5010:
  InvalidParam passed to GetItem(id)` the first time a shape created a sync
  earlier is touched again. Both experiments are queued in the batch that
  creates their shape now, which is also the shape a real merge has. They draw
  on a slide the probe inserted, so the positional sweep removes them, and they
  do not draw at all when no insert landed.
- The reader scored question four against `"Hello Ada here and 1-2".replace("2",
  "BBB")`, a string neither offset model produces. Both experiments' text and
  both predictions are single-sourced now, and a guard checks each prediction is
  what applying that model by hand actually gives.

- The probe reported a verdict on a question it had not asked. The tag read
  lands on the last slide in the deck, which is the inserted one only if the
  insert worked; the first real sheet had every insert throw, so that read fell
  on the user's own title slide and the reader announced that the metadata
  scheme needed rethinking. The scheme had never been tested. `tagVerdict` now
  answers "not asked" whenever the slide carrying the tag did not land.
- The probe's fixture deck carried `<a:themeElements/>`, which is
  schema-invalid: `CT_BaseStyles` requires clrScheme, fontScheme and fmtScheme,
  all three mandatory, and `KeepSourceFormatting` is precisely the path that has
  to import the source theme. Checked against a deck PowerPoint itself accepts.
  The fixture also gained the docProps parts and a real `xfrm` on its text box,
  so a rejection has fewer places to hide.
- Four separate host calls in the substring probe shared one `catch`, so a throw
  named none of them and the sheet said `InvalidArgument` about a statement
  nobody could identify. Each call now stamps what it is doing, and the error
  carries it.

### Added

- A control arm on question one, and the rule that reads it. Inserting the
  presentation's own bytes asks a question only one of the two readings of
  `InvalidArgument` survives: a deck PowerPoint saved seconds ago cannot be a
  malformed package, so a host that refuses it is refusing insertion itself.
  `insertionBlame` says OURS, THE HOST, or CANNOT TELL, and never guesses.
- A second insert of the same package under `UseDestinationTheme`, which does
  not import the source theme where `KeepSourceFormatting` must.
- The generated snippet is typechecked against the real Office.js types in CI.
  It sits outside tsconfig's include and is pasted into an editor that runs it
  before anyone reads it, so nothing else would catch a misspelled option key.

- The host probe: a Script Lab snippet that asks the four questions only a real
  PowerPoint can answer, and a reader that interprets the sheet it produces.
  The snippet collects observations and makes no judgements; every reading
  happens in the repository, where it is tested.
- The merge plan: a block of contiguous slides repeated once per record,
  record-major so a record's slides stay together, with conditional slides
  skipped in place and a policy for what an empty cell means.
- The runner that carries a plan out against a package: clone the template
  slide, merge the copy, tag the copy. The template is never touched, so it can
  be merged again.
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
