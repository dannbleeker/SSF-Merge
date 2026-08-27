# SSF Merge — manual

Mail merge for PowerPoint. You build one slide, or one block of slides, put
placeholders where the data goes, and SSF Merge produces one copy per row.

> **Status.** The engine is built and tested; the task pane is not written yet.
> Sections marked *planned* describe behaviour that is designed and agreed but
> not yet shipped. Nothing here is aspirational: a section moves out of
> *planned* in the same change that makes it true.

## Contents

- [The idea](#the-idea)
- [Placeholders](#placeholders)
- [Formats](#formats)
- [What repeats](#what-repeats)
- [Your data](#your-data)
- [Where the merged slides go](#where-the-merged-slides-go) *(planned)*
- [What happens to the template](#what-happens-to-the-template) *(planned)*
- [The pane](#the-pane) *(partly planned)*
- [Installing it](#installing-it)
- [Tags SSF Merge writes](#tags-ssf-merge-writes)
- [Limits](#limits)

## The idea

A .pptx file is a zip of XML parts. SSF Merge does the whole merge inside that
file and hands PowerPoint the finished deck in one operation. Your formatting
survives because nothing re-authors it: each piece of text keeps the exact run
properties the designer gave it, and only the characters change.

## Placeholders

Write a field where the value should appear:

```
{{Name}}
```

Field names may contain letters, digits, underscores and dots, so `{{Customer.Name}}`
works if your columns are named that way.

**A placeholder split across formatting still works.** PowerPoint constantly
stores `{{FirstName}}` as `{{Fir` + `stName}}` after an edit or a spellcheck
pass. SSF Merge matches against the whole paragraph, not one piece at a time.

**The value takes the formatting of the run the placeholder starts in.** If you
want the merged name bold, make `{{` bold. This is the one rule worth knowing
when a merged value comes out looking wrong.

**A field with no matching column is left visible.** You will see
`{{Territory}}` on the slide rather than a blank space. A blank slide looks
finished; a visible placeholder does not, which is the point.

A field name may be written in any language — `{{Beløb}}`, `{{Größe}}`,
`{{Πλήθος}}` — and matches the column header exactly as the sheet spells it.
Letters, digits, underscores and dots; a name made only of spaces or punctuation
is not a field.

## Formats

Add a format after a pipe:

```
{{Revenue|number:2}}      1 234 567,89
{{Revenue|number}}        1 234 568
{{Start|date:dd MMM yyyy}} 01 Mar 2026
{{Start|date}}            01-03-2026
{{Region|upper}}          NORDICS
{{Region|lower}}          nordics
```

| Format | Argument | Notes |
| --- | --- | --- |
| `number` | decimal places, default 0 | Space for thousands, comma for decimals |
| `date` | a pattern, default `dd-MM-yyyy` | `yyyy` `yy` `MMM` `MM` `dd` `d` |
| `upper` | none | Locale-aware, so `måned` becomes `MÅNED` |
| `lower` | none | Locale-aware |

**A value that does not match its format is printed unchanged.** `{{Note|number}}`
on the text "n/a" gives you "n/a", not a blank and not an error marker. The cell
is what you typed, and showing it is more useful than hiding it.

### Numbers

Both European and American forms are read. `1,5` is one and a half; `1,500` is
one thousand five hundred; `1.234,56` and `1,234.56` are the same number.

### Dates

`2026-03-01` and `3 March 2026` are read. So is `15/01/2026`, because only one
reading fits.

**`03/01/2026` is refused.** It is 3 January in Copenhagen and 1 March in
Chicago, and nothing in the cell says which. The value is printed as it stands
rather than guessed at. A deck that draws perfectly and is two months wrong is
the worse outcome. Write the date as `2026-03-01` in your source if you want it
merged as a date.

## What repeats

A template is one or more **contiguous** slides, marked as a block. Three slides
per customer is the ordinary case. Step 1 of the pane is where you name it: the
number of the first slide and the number of the last, as the thumbnail rail
counts them.

Picking the block by *clicking* the slides is still planned. Two boxes are what
is built, and they are checked as you type. A block that ends before it starts,
a slide 0 or a fraction is **refused**, in a sentence quoting what you typed,
before a template read is spent on it.

A block running past the end of the deck is only a **warning**, and the button
stays live. The pane counts your deck when it opens and again each time you
press "Use slides N to M"; between those it can be out of date, and refusing a
block because of a count taken ten minutes ago would tell you a slide you can
see does not exist. The real check happens a moment later against the slide list
PowerPoint answers with at that instant.

- The deck's own order is the order each record gets.
- Slides must sit next to each other. Reorder them in the thumbnail pane first.
- Records are emitted whole: all of record 1's slides, then all of record 2's.
- A slide can be conditional, so a record gets two slides or three. Its position
  never changes; it is skipped in place.
- A condition names a column. The slide is emitted when that column's cell has
  content. **A blank cell is false, and so are the words `false`, `falsk`, `no`,
  `nej`, `off` and `0`**, whatever their capitalisation. That short list exists
  because a spreadsheet boolean does not arrive as a boolean: Excel writes it
  out as a localised word, and treating `FALSK` as content would emit every
  slide it was told to leave out. Anything not on the list is content.
- A condition naming a column your data does not have **emits the slide anyway**
  and reports the problem. Dropping it would hide an authoring mistake behind
  output that looks finished.
- One deck can hold several blocks over different data. Slides 4 to 6 over
  customers while slides 9 to 10 repeat over products is an ordinary report.

## Your data

*Partly planned.* Pasting works; the other two sources are still to come.

| Source | State |
| --- | --- |
| Paste a range copied from Excel | built — step 2 of the pane |
| A .csv or .xlsx file | planned |
| An Excel table on OneDrive or SharePoint via Microsoft Graph | planned |

Select the range in Excel, copy, and paste into the box on step 2. That arrives
tab-separated and is read as such; a comma-separated paste is read too. The
columns found are listed under the box beside the row count, deliberately — a
copy that came through as plain text parses into ONE column, and a row count on
its own looks perfectly healthy when that happens.

A header row with nothing under it is refused rather than counted as zero rows,
because "Add 0 slides" does not say which half was wrong.

The first row is the header. A column with no header is named `Column 3` rather
than dropped, and two columns with the same header become `Name` and `Name 2`,
because silently losing a column is worse than an ugly name. Rows that are
entirely blank are skipped.

Each column's type is detected from its values: a column is a number only if
every filled cell is one, and a date only if every filled cell is an
unambiguous date. One "n/a" makes the column text, which is the safe answer.

## Where the merged slides go

*Planned.* Three choices:

1. **Into this deck**, after the template block.
2. **Into a new presentation**, which opens beside the current one. Better for
   large merges, because a very long deck is slow to edit.
3. **One file per row**, saved to OneDrive. Requires signing in.

## What happens to the template

*Planned.* When merging into the same deck, the template block can:

- **stay where it is** (default) — you can run the merge again later;
- **move to the end** — out of the way, still re-runnable. PowerPoint's add-in
  API cannot hide a slide, so this is as close as it gets;
- **be deleted** — one way. Deleting ends re-run for that deck, and it happens
  only after the merge is confirmed to have landed.

## Tags SSF Merge writes

Merged slides carry metadata inside the file. You never need to touch it, but it
is what makes undo and re-run possible, and it is visible to any tool that reads
PowerPoint tags.

| Tag | On | Meaning |
| --- | --- | --- |
| `SSF_MERGE_RUN` | a merged slide | Which merge produced it |
| `SSF_MERGE_BLOCK` | a template or merged slide | Which template block it belongs to |
| `SSF_MERGE_SEQ` | a template or merged slide | Its position within the block |
| `SSF_MERGE_RECORD` | a merged slide | Which row it was made from |

Tags that other tools wrote are kept. SSF Merge merges its own keys into an
existing tag list rather than replacing it.

## The pane

Four steps, and the step number is shown throughout — "Step 2 of 4" states how
much is left, which is the question a first-time user actually has.

| step | what you do |
| --- | --- |
| 1 · Template | Name the first and last slide of the set that repeats. They must be next to each other. |
| 2 · Fields | See the placeholders found in those slides, and which ones have no column behind them. |
| 3 · Preview | Merge the first row into your deck so you can look at it, then remove it. |
| 4 · Merge | Add the slides, with the count on the button. |

Three things about it are deliberate and worth knowing.

**Exactly one filled button per screen, always last, and it names what it does
with the number in it.** "Add 720 slides" is a statement you can check against
the deck in front of you; "Merge" is only a promise.

**Every slide is named by the number in the thumbnail rail.** Never an id — this
host refuses ids for slides a run has just added — and never a zero-based index,
which is a number you have no way to see.

**One orange thing per screen.** Normally the small tick above the heading. When
something is temporarily untrue on the slide — a preview showing, a placeholder
with no column — the orange moves there and the tick goes away. Two oranges in
one glance and neither means anything.

**A link back on every screen but the first.** A wizard you can only walk
forwards through is one you restart to change a number.

**One host call at a time.** While the pane is reading your slides or merging,
the button says so and nothing is pressable — including after going back a step
and forward again, which is the way a wizard usually lets you start the same
long job twice. Once a merge has added its slides the button says how many and
stays down; change the block or the data and it arms again.

All four steps are built and reachable from the screen: name the block, paste
the rows, look at one, press the button. Pressing "Use slides N to M" reads
those slides out of the open deck and lists the placeholders it actually found
in them, so step 2 shows the real fields rather than a guess.

### What a preview actually is

**It merges.** Pressing "Preview the first row" runs the ordinary merge over
your first row and puts the result at the end of your deck, exactly as the full
merge would. "Remove the preview" deletes those slides again.

That is the point rather than an implementation detail. What you are looking at
was produced by the code that will produce the other 239 slides, so if it looks
right, the rest will be right. A preview drawn some other way is a preview of
something nobody is going to get.

**Your template is never touched.** The obvious way to build this — write the
row's values onto your template slide and put the original text back afterwards
— would have damaged the one thing this add-in exists to protect. Setting a
shape's text through the PowerPoint API re-authors it, so the text would come
back and the formatting would not: silently, on the master copy every merged
slide is cloned from.

Two consequences worth knowing. The preview lands at the **end of the deck**,
not next to your template, because that is the one insertion point this add-in
has tested on a real host. And if you close the pane while a preview is showing,
those slides simply stay — they are ordinary slides and you can delete them.
The card names them for that reason.

The pane follows PowerPoint's theme, read once when it opens. There is no
theme-change event for a PowerPoint task pane — the one in the Office typings
belongs to Outlook — so switching PowerPoint between light and dark while the
pane is open needs the pane reopened.

## Installing it

The add-in is a **manifest** — a small file naming a web page — plus the page
itself, which is hosted at
<https://ssf-merge.struktureretsundfornuft.dk>. Nothing is installed onto
your machine.

Download **`manifest-prod.xml`** from the
[latest release](https://github.com/dannbleeker/SSF-Merge/releases/latest) and
sideload it:

| where | how |
| --- | --- |
| PowerPoint on the web | Home → Add-ins → More Add-ins → **My Add-ins** → Upload My Add-in, and pick the file |
| PowerPoint on Windows | Put the file in a folder, share the folder, then File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs and add the share. Restart PowerPoint; the add-in appears under **Shared Folder** |
| PowerPoint on Mac | Copy the file to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef` |
| A whole tenant | An administrator deploys **`manifest-prod.json`** (the unified manifest) from the Microsoft 365 admin centre |

Then look on the **Home** tab for the **Mail merge** button.

### What needs re-installing, and what does not

Almost nothing. The pane is served from the web, so a change to the merge, the
steps, the wording or the layout reaches you the next time you open the pane —
no re-install, nothing to download.

The **manifest** is the exception. Re-sideload only when the manifest itself
changes: the ribbon button, the permissions, the display name, the icons, or
the page the button opens. Those changes are called out in the changelog.

### Why it does not say which PowerPoint it needs

The manifest declares no `<Requirements>` block, deliberately. SSF Merge needs
**PowerPointApi 1.3** — merge metadata lives in slide tags, which arrived in 1.3
— and that is checked when the pane opens, not declared in the manifest.

A declared requirement set that your PowerPoint does not meet makes the add-in
**vanish from the ribbon** with no message at all: nothing to see, nothing to
report, nothing to search for. The runtime check can tell you which version is
missing and what it costs you, which is worth more than a silent absence.

**The cost of that choice, stated plainly.** Because the manifest declares
nothing, Microsoft's own validator reports the add-in as installable on every
PowerPoint back to **2013 on Windows** — and PowerPoint 2013 does not have
PowerPointApi 1.3. So on an old enough PowerPoint the add-in installs, appears
on the Home tab, opens, and then tells you it cannot run and why. That is the
trade: a message you can read and act on, instead of a button that was never
there.

## When a cell is not what its format expects

The rule throughout is that a cell the engine cannot read is **printed as it
stands**, never guessed at and never blanked. A merged deck that draws perfectly
and is two months wrong is worse than one showing the cell untouched, because
nobody checks the first one.

So: an impossible date — `29/02/2025`, `31/04/2026`, `13/13/2026` — is left
alone rather than rolled forward into the real date that follows it. An
ambiguous slash date like `03/01/2026` is left alone for the same reason. A
`number:` format asking for impossible decimals leaves the cell alone rather
than failing the merge.

Dates are read and written in the same zone, so `1 Mar 2026` prints as
`01 Mar 2026` wherever you are.

## Limits

- **Charts and SmartArt are not merged.** Their text lives in separate parts
  with their own embedded workbooks. Placeholders inside them are left alone.
- **Cut and paste on PowerPoint for the web loses shape tags**
  ([office-js#3784](https://github.com/OfficeDev/office-js/issues/3784)). A
  merged slide cut and pasted into another deck loses its run tag, so undo will
  no longer find it.
- **A very long deck is slow to edit**, which is PowerPoint's behaviour and not
  something an add-in can fix. Merging into a separate presentation avoids it.
