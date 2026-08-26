# SSF Merge

Mail merge for PowerPoint. One template block of slides, one set of slides per
row of your data, and formatting that survives because nothing re-authors it.

Part of the SSF add-in family from
[StruktureretSundFornuft](https://struktureretsundfornuft.dk).

## Status

Early. The engine's spine is here and tested; the task pane and the host layer
are not written yet.

| Piece | State |
| --- | --- |
| Package layer (`Pkg`, parts, rels, content types, slide ids) | done |
| Slide cloning, with fresh creation ids and per-copy notes pages | done |
| Tags written into the file (`ppt/tags/tagN.xml`) | done |
| Run-aware text replacement | done |
| Data parsing, type detection, formatting | done |
| Merge plan (blocks, records, conditional slides) | next |
| Office.js host layer, task pane, manifests | not started |

## How it works

A .pptx is a zip of XML parts. The merge happens **in the file**, not through
the PowerPoint API, and the finished deck is handed to PowerPoint in a single
`insertSlidesFromBase64` call.

That is not a preference. Setting a shape's text through Office.js re-authors
it, and [office-js#5858](https://github.com/officedev/office-js/issues/5858)
reports custom bullets reverting to defaults when it does. In the file, each
`<a:r>` run keeps its own `<a:rPr>` and only the characters change, so a
designer's slide comes out the way they built it.

```
data (paste, .xlsx, Graph)  ─┐
                             ├─► merge engine (pure TypeScript) ─► .pptx bytes ─► one host call
template block (the deck)   ─┘
```

The engine imports nothing from Office.js, and a test enforces that. It runs in
the task pane, in a Node CLI and in the suite with no PowerPoint anywhere.

## Template syntax

`{{Field}}` for a value, `{{Field|format}}` for a formatted one:

```
{{Name}}
{{Revenue|number:2}}
{{Start|date:dd MMM yyyy}}
{{Region|upper}}
```

Placeholders survive being split across runs, which PowerPoint does constantly:
`{{Fir` + `stName}}` is matched and replaced as one field. The value inherits
the formatting of the run the placeholder **starts** in, so format the opening
brace the way you want the merged value to look.

A field with no matching column is left visible rather than blanked. A blank
slide looks finished; `{{Territory}}` does not, and that is the point.

## Documentation

| Document | What it is for |
| --- | --- |
| [docs/MANUAL.md](docs/MANUAL.md) | How to use it: placeholders, formats, blocks, data, limits |
| [docs/BACKLOG.md](docs/BACKLOG.md) | What is open, what it would cost, and what has been rejected |
| [CHANGELOG.md](CHANGELOG.md) | What changed, newest first |
| [CLAUDE.md](CLAUDE.md) | Project memory: architecture, host rules, conventions |

These are kept in step with the code by `test/docs.test.ts`, which reads the
formats and tag keys out of the source and fails when the manual has not caught
up. A feature and its documentation land in the same change or neither does.

## Commands

```bash
npm install
npm test          # the suite
npm run typecheck
```

## Licence

MIT.
