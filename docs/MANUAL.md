# SSF Merge — manual

Mail merge for PowerPoint. You build one slide, or one block of slides, put
placeholders where the data goes, and SSF Merge produces one copy per row.

> **Status.** The engine, the task pane and the four steps are built and
> shipped. What is still *planned* is marked, and it is now single options
> inside sections rather than whole sections: sending the merge somewhere other
> than this deck, doing anything to the template afterwards, and reading data
> from a file or from Excel.
>
> Nothing here is aspirational: a line moves out of *planned* in the same change
> that makes it true. This block said the pane was not written for some days
> after it shipped, which is the failure it exists to prevent.

## Contents

- [Your first merge](#your-first-merge)
- [The idea](#the-idea)
- [Placeholders](#placeholders)
- [Formats](#formats)
- [What repeats](#what-repeats)
- [Your data](#your-data)
  - [Choosing which slides are conditional](#choosing-which-slides-are-conditional)
- [Where the merged slides go](#where-the-merged-slides-go) *(one of three)*
- [What happens to the template](#what-happens-to-the-template) *(default only)*
- [The pane](#the-pane)
- [Installing it](#installing-it)
- [Tags SSF Merge writes](#tags-ssf-merge-writes)
- [Limits](#limits)

## Your first merge

Five minutes, on a deck you do not mind adding slides to. Nothing here deletes
anything: a merge only ever **adds** slides after your template, and step 4
offers to take them straight back out again.

1. **Make a template slide.** Any slide will do. Put two placeholders on it,
   typed exactly like this, braces included:

   ```
   {{First}}
   {{City}}
   ```

   Note which slide number it is in the thumbnail rail on the left — say it is
   slide 3. A template can be several slides in a row; one is enough to start.

2. **Open the pane.** Home tab → **Mail merge**.

3. **Step 1 — Template.** Type `3` into both boxes. The button reads **Choose
   the slides that repeat** until the numbers make sense, then becomes **Use
   slides 3 to 3** — press it. The pane reads the slide and says it found 2
   placeholders. If it says 0, the braces are probably curly quotes, or the
   field name has a space in it.

4. **Step 2 — Fields.** Paste this into the box, including the header row:

   ```
   First	City
   Ada	London
   Grace	New York
   Katherine	Hampton
   ```

   Copy it out of a spreadsheet, or type it with real tab characters between the
   columns. The pane will say **3 rows · First, City** and show a chip per
   placeholder, with any chip that has no column behind it outlined. Press
   **Use 3 rows**.

5. **Step 3 — Preview.** Press **Preview the first row**. One row is merged
   into your deck so you can look at it, and **Remove the preview** takes it
   out again. This is an ordinary one-row merge, not a mock-up — what you see
   is what step 4 produces. **Skip to the merge** goes on without previewing.

6. **Step 4 — Merge.** The button reads **Add 3 slides** — the count, not the
   word "merge" — and the line above says where they land and how big the deck
   will be afterwards. Press it, look at the result, and if you do not like it
   press **Remove these slides**.

If any step does anything other than this, the pane's run record is at the
bottom of the screen once a run finishes — open **What this run did, step by
step**, select it, and copy it. It carries no cell values, only structure.

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

Select the slides in the thumbnail rail and press **Use the slides I have
selected**, or type the two numbers. Either way they are checked as you type.

A selection with a **gap** in it is refused rather than closed up: a template
block is slides that repeat *together*, in order, so quietly filling the gap
would add a slide to every row that you never picked. A block that ends before it starts,
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
  never changes; it is skipped in place. Set it on **step 2**, under *Every
  slide, every row* — one dropdown per slide in your block, offering the columns
  in the data you attached.
- A condition names a column. The slide is emitted when that column's cell has
  content. **A blank cell is false, and so are the words `false`, `falsk`, `no`,
  `nej`, `off` and `0`**, whatever their capitalisation. That short list exists
  because a spreadsheet boolean does not arrive as a boolean: Excel writes it
  out as a localised word, and treating `FALSK` as content would emit every
  slide it was told to leave out. Anything not on the list is content.
- A condition naming a column your data does not have **emits the slide anyway**
  and reports the problem — under the dropdowns before you merge, and in the
  sentence after it. Dropping it would hide an authoring mistake behind output
  that looks finished. You can reach this without typing a name: pick a column,
  then paste data that does not have it.
- Conditions belong to the **template**, so a new paste keeps them. They are
  keyed by slide number, so **changing the block clears them** — "slide 5 only
  when Renewal" is about the fifth slide of the deck, and a block that starts
  somewhere else would silently apply it to a different slide.
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

### Choosing which slides are conditional

Under the placeholder chips is a line reading *Every slide, every row*. Opening
it shows one dropdown per slide in your block, each offering the columns in the
data you attached: leave it on **Always**, or pick **Only when [column]**.

The columns are offered rather than typed, because the engine matches a
condition against a column name exactly and a typed name that does not match is
a silent no-op you would find by counting slides in the output.

Shut, the line says how many slides are conditional, so the answer is on screen
without opening anything.

### Choosing which rows merge

Every pasted row is in by default. Under the paste box is a line saying how many
rows are in and how many are out; opening it shows a checkbox per row, labelled
with that row's first column, and a search box above them.

Search matches any cell in a row, not only the first column. It changes what the
list SHOWS, never what is ticked — so you can search for a region, untick the
three rows you do not want, clear the search, and those three stay out.

The list shows at most 60 rows at a time and says how many it did not show.
Search to reach the rest.

The merge button, the summary and the preview all count the ticked rows. Untick
everything and the button says so instead of merging nothing. Pasting new data
clears the filter, because a row number means nothing against different data.

## Where the merged slides go

**Into this deck, after the template block.** That is what ships, it is not a
setting, and the merge screen says where the slides will land and how large the
deck will be afterwards before you press anything.

Two more destinations are designed and not built:

- **Into a new presentation**, which would open beside the current one. Better
  for large merges, because a very long deck is slow to edit. *Planned.*
- **One file per row**, saved to OneDrive. *Planned*, and blocked rather than
  merely unbuilt: a task pane cannot hand you a downloaded file
  ([office-js#1511](https://github.com/OfficeDev/office-js/issues/1511)), so it
  needs an upload-and-link route.

## What happens to the template

**Nothing. It stays exactly where it is**, which is what ships and what makes a
merge re-runnable: change a row, run it again. A merge only ever ADDS slides.

Two alternatives are designed and not built, both of them one-way enough to be
worth doing carefully rather than quickly:

- **move it to the end** — out of the way, still re-runnable. PowerPoint's
  add-in API cannot hide a slide, so this is as close as it gets. *Planned.*
- **delete it** — ends re-run for that deck, so it would happen only after the
  merge is confirmed to have landed. *Planned.*

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

## If PowerPoint takes only part of the merge

It reports in ROWS, not slides, because rows are what you pasted:

> PowerPoint took only part of the merge: 2 of 3 rows landed complete; row 3 got
> 1 of its 2 slides. Take the slides back and run it again.

That last sentence is the advice. A row with some of its slides is worse than a
row with none — it looks finished, and every row after it looks correct too, so
a short merge is easy to miss until somebody reads slide 141.

**Undo takes back everything that landed**, including the incomplete row. It
removes the slides the deck actually gained, not the number the merge was
aiming for, so there is nothing left stranded.

## Taking a merge back

After a merge lands, the pane offers **"Remove slides 13 to 732, which this
merge added"**. It names the slides rather than saying "undo", because it is
deleting part of your presentation and you should be able to check before
pressing.

It works by POSITION, not by looking the slides up: a slide a run has just
added cannot be found by id on PowerPoint for the web. The sweep is clamped so
it can never reach an index below the deck's size when the merge started, so
nothing you had before the run can be touched.

**If your deck grew after the merge, it refuses.** Add slides yourself, or have
a co-author add some, and the last slides in the deck are no longer the ones the
merge added — so the sweep says nothing was removed and leaves the deck alone.
Delete those slides by hand. This is the safe direction: the alternative is
deleting somebody else's work.

A sweep that removed only some of them keeps the offer up, because the rest are
still there and only you can finish.

## When something goes wrong

**The pane says what it is waiting on.** A merge is legitimately quiet for a
while: reading the template is allowed ninety seconds and handing the package to
PowerPoint another sixty, so up to two and a half minutes can pass with nothing
happening on screen. While it waits the pane names the call — "Waiting on
PowerPoint: inserting the merged deck…" — so you can tell a slow step from a
stuck one, and so you can say which step it stopped at.

**Every run leaves a record.** Under the outcome there is a collapsed *"What
this run did, step by step"*. Open it and you get every call the run made, in
order, with how long each took:

```
=== SSF MERGE RUN LOG ===
   0.0s  host  issued    call=counting the deck's slides budget=15000
   0.1s  host  answered  call=counting the deck's slides ms=94
   0.1s  host  issued    call=exporting the template slides budget=90000
   2.4s  host  answered  call=exporting the template slides ms=2311
   2.4s  host  issued    call=inserting the merged deck budget=60000
  14.9s  host  answered  call=inserting the merged deck ms=12470
=== END ===
```

Select it and copy it. That is the whole channel: a task pane cannot open
developer tools and cannot save a file, so the text on screen is how the record
reaches anybody. If a merge goes wrong, this is the thing worth sending.

A call that never came back shows its `issued` line with no `answered` after it,
which names the step that stopped.

**"No placeholders were filled."** The merge added the slides and changed
nothing on them, which almost always means the placeholder names in your
template do not match your column headers. Check the spelling and the braces —
`{{First name}}` matches a column headed `First name`, and nothing else does.

## Installing it

The add-in is a **manifest** — a small file naming a web page — plus the page
itself, which is hosted at
<https://ssf-merge.struktureretsundfornuft.dk>. Nothing is installed onto
your machine.

Get **`manifest-prod.xml`** and sideload it. Two places to get it, and the
second one only once a version has been cut:

- [the file on `main`](https://github.com/dannbleeker/SSF-Merge/raw/main/manifest-prod.xml)
  — always current, and the right one while the add-in is still being tested;
- the [latest release](https://github.com/dannbleeker/SSF-Merge/releases/latest),
  once there is one. **There is none yet**, so that link is a 404 today; this
  manual said to use it before anybody checked.

Both name the same hosted page, so the pane you get is the same either way — the
manifest is only a pointer, and it is the pointer that is versioned, not the
add-in.

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
**PowerPointApi 1.2** — reading the deck, inserting the merged slides and taking
them back again — and that is checked when the pane opens, not declared in the
manifest.

Two things sit above that floor and are simply absent on an older PowerPoint
rather than blocking it: reading just the template slides instead of the whole
file (1.10), and the **Use the slides I have selected** shortcut (1.5). Typing
the two slide numbers works everywhere.

A declared requirement set that your PowerPoint does not meet makes the add-in
**vanish from the ribbon** with no message at all: nothing to see, nothing to
report, nothing to search for. The runtime check can tell you which version is
missing and what it costs you, which is worth more than a silent absence.

**The cost of that choice, stated plainly.** Because the manifest declares
nothing, Microsoft's own validator reports the add-in as installable on every
PowerPoint back to **2013 on Windows** — and PowerPoint 2013 does not have
PowerPointApi 1.2. So on an old enough PowerPoint the add-in installs, appears
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
