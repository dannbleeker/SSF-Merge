# The Marketplace listing

Everything Partner Center asks for, written here first so it is reviewable,
diffable and not retyped into a web form from memory.

Lengths are Microsoft's, from
[Create effective listings](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/create-effective-office-store-listings):

| Field | Maximum | Recommended | Key message in the |
| --- | --- | --- | --- |
| Name | 50 characters | 30 | first 30 characters |
| Summary | 100 characters | 70 | first 30 characters |
| Description | 10,000 characters | 300–500 words | first 300 words |

That page states the description maximum as 10,000 in its table and 4,000 in its
prose. The text below is well under 4,000, so the contradiction costs nothing.

## Name — a decision, not a draft

The manifest says **SSF Merge**, and the Partner Center name has to match it.
Microsoft's own naming guidance argues against it on two counts:

> Avoid acronyms that might be unfamiliar to potential users.

> Make the purpose or benefits of your solution clear. Don't rely on your brand
> to communicate what your solution does.

"SSF" is an acronym a stranger cannot expand, and "Merge" alone does not say
merge *what*. Their suggested pattern is **Function + for + brand**, with one
exclusion that rules out the obvious answer:

> Don't include the Microsoft product name because it already appears on your
> landing page, and in Microsoft Marketplace and in-product Store search results.

So "Mail Merge for PowerPoint" is out.

**Suggested: `Slide Mail Merge`** — 16 characters, says what it does, no
acronym, no Microsoft product name, and it carries the words somebody would
actually search for.

This is left as a suggestion rather than applied. It is a product rename: the
manifest's `DisplayName` has to change with it, and that is the name users see
in the ribbon after install. It is the owner's call, not a formatting fix.

## Summary

Maximum 100, recommended 70. Partner Center asks for "a single sentence", plain
text, no line breaks, and says it appears in search results and summary views.

```
Paste a table and get one set of slides per row, formatting intact.
```

67 characters. The verb phrase is the description's own — "Mark the slides that
repeat, paste a table, and get one set of slides per row" — so the two fields
say the same thing in the same words rather than drifting into synonyms. It
leads with the outcome rather than the name, which is what the guidance asks
for: the name is already on the tile beside it, and the first 30 characters read
"Paste a table and get one set".

The previous draft, "One set of slides per row of your table, with fonts and
layout intact.", was a noun phrase rather than a sentence, and "fonts and
layout" was a stand-in for the word the rest of the product uses, which is
"formatting".

**This depends on the name.** It works when the offer name carries the words
somebody searches for. If the name becomes `Slide Mail Merge`, "mail merge" is
on the tile and the summary is free to be the outcome. If it stays `SSF Merge`,
the summary is the only place those words appear, and this is the one to use
instead:

```
Mail merge for slides: paste a table and get one set of slides per row, formatting intact.
```

90 characters, inside the limit, at the cost of spending the first 22 on the
category rather than on what the thing does.

## Description

Maximum 10,000 characters in Microsoft's table and 4,000 in its prose on the
same page; 300 to 500 words recommended, key message in the first 300. This one
is 378 words and 2,293 characters including its tags, so the contradiction costs
nothing either way.

**The field takes HTML, not markdown.** Partner Center's own help says "simple
HTML tags", and Microsoft publishes the list:
[supported HTML tags for offer descriptions](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/supported-html-tags).
Supported are `b`, `i`, `br`, `p`, `ul`/`li`, `ol`/`li` and `h1` through `h6`.
There is no `code` and no `a`, so a placeholder cannot be marked up as code and
a link cannot be embedded. An earlier draft of this section was markdown, which
would have put literal asterisks and backticks in the listing.

It opens on the manifest's own `Description`, because
[the guidance](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/create-effective-office-store-listings)
says the listing description "should match the description in your manifest as
closely as possible". The same page says to include search keywords, which is
why Excel is named: pasting a range copied from Excel is step 2 of the pane.

It answers the four questions that page asks a description to answer: how it
benefits the user, what is special about it, the different ways it could be
used, and who would use it.

Paste this as it is, and use Partner Center's preview before saving.

```html
<p>Mail merge for PowerPoint. Mark a block of slides as your template, paste a table, and get one set of slides per row, each with that row's name, numbers, dates and picture in place. Your rows can be a range copied straight out of Excel, or any table you can paste.</p>
<p>What used to take an afternoon of duplicating a slide and editing each copy takes a minute, and every name and number comes from your table rather than from your typing.</p>
<p><b>Your formatting survives, because nothing re-authors it.</b></p>
<p>Most tools that fill in a slide rebuild the text, and your fonts, spacing and theme colours quietly change. This one copies the slides you designed and replaces only the placeholder, leaving every other property exactly as PowerPoint wrote it. A deck merged this way looks like the deck you built.</p>
<h2>What it fills in</h2>
<ul>
<li>Text anywhere on the slide, including speaker notes</li>
<li>Numbers and dates, formatted the way you ask: {{Revenue|number:0}}, {{Renewal|date:d MMM yyyy}}</li>
<li>Pictures, one per row, cropped to fill the frame you drew without distortion</li>
<li>Chart data and titles, including the workbook behind a chart, so Edit Data shows that row's figures</li>
<li>SmartArt, in both halves it is stored in, so the diagram on screen says the right thing</li>
</ul>
<h2>Before it does anything</h2>
<p>It tells you what it is about to do, how many slides it will add and where they will land, and it shows you one row merged so you can look at a real result before you commit to the rest. Afterwards it will take those slides back out again, and it names exactly which ones it will remove.</p>
<h2>Who it is for</h2>
<p>Anyone who makes the same deck repeatedly with different names on it: account reviews, candidate packs, store or regional reports, class or cohort summaries, certificates, starter packs. If you have ever duplicated a slide twelve times and edited each copy by hand, this replaces that afternoon.</p>
<h2>Where your data goes</h2>
<p>Nowhere. There is no server, no account and no sign-in, and the add-in makes no network requests while it runs. Your slides and your table are read inside the task pane, on your own device, and the merged slides are written straight back into the presentation you already have open.</p>
```

## Search keywords

Up to three. Partner Center's own help: "Don't add words or acronyms that are
already included in your product's name, summary, or description." Nothing in
Microsoft's published documentation covers this field, so that sentence is the
whole rule, and it is the one thing here a machine can check. `listing.test.ts`
does, against the manifest's `DisplayName` and the two blocks above.

```
bulk
personalized
csv
```

Each is absent from the name, the summary and the description, so no slot is
spent repeating a word the search engine already has from those.

- **bulk** is what somebody types when they have been doing this by hand. It is
  the highest-intent word not already spent.
- **personalized** is the standard mail-merge word for the outcome, and the
  description never uses it. American spelling deliberately: it is the one with
  the search volume, and the store's default locale is en-US.
- **csv** is true rather than aspirational. `src/core/data/recordset.ts` reads
  `["\t", ",", ";"]`, so tab, comma and semicolon pastes all parse.

**The caveat on `csv`.** The add-in reads a pasted table; there is no file
picker, and the manual records that as a decision rather than a gap. Somebody
searching "csv" may be after file import and will find paste-only. Pasting the
contents of a .csv works, so the keyword is accurate, but `automation` is the
swap if that expectation is not worth inviting. It is free and safe and generic
enough to compete with half the store, which is the trade.

`duplicate` was considered and rejected: the description says "duplicating a
slide", and any stemmer treats those as the same word.

## Screenshots

Required: at least one, at **1366×768**, under 1024 KB, PNG/JPEG/GIF, up to five.

`shots/` holds them, captured from the demo deck below in one pass through the
pane on PowerPoint for the web. `test/listing.test.ts` opens every one on every
run and fails on the wrong size, a truncated file, a gap in the numbering, or
the reappearance of the two placeholders that used to sit here.

Those placeholders — `01-attach-your-rows.png` and `02-see-what-it-will-add.png`
— were real captures of the real product and were still wrong to upload, for
five reasons worth keeping written down, because every one of them is a trap the
next capture session can walk back into:

1. The photograph was the kit's crop fixture, a blue rectangle with four yellow
   dots positioned to prove a cover-crop keeps the right axis. A customer reads
   it as a broken image.
2. The chart values were 18 and 42, which are two numbers chosen to tell a swap
   from a fill.
3. The title bar read `SSF-Merge-test-template`.
4. Slide 1 was a wall of instructions addressed to whoever ran the test round.
5. **There was a profile photo of a real person in the top right corner.**
   Microsoft: "Be sure to remove any personal information from your images that
   you don't want customers to see."

### The deck they are shot from

`demo/` holds one, built rather than improvised so the numbers in the pictures
agree with each other:

| File | What it is |
| --- | --- |
| `demo/build-demo.py` | writes the two below; run by hand, like `test-kit/build-template.py` |
| `demo/Quarterly-business-review.pptx` | a cover slide and a two-slide template block |
| `demo/rows.txt` | three accounts, four columns |

Three rows and a two-slide block, so the merge adds **six slides** and the deck
goes from three to nine. `test/listing-demo.test.ts` merges it on every run and
fails if any placeholder has no column behind it, because an unmatched field is
deliberately left on the slide by the engine — right in the product, and a
screenshot with `{{Account}}` printed on it here.

It has no picture placeholder, and that is a decision rather than an oversight.
Merging pictures is real and it is in the description, but a screenshot of it
needs a photograph that looks like one, and anything generated would be a
coloured rectangle: the placeholder problem again in a different colour.

### The five

| File | What it shows |
| --- | --- |
| `shots/1-mark-the-block.png` | slides 2 and 3 selected in the rail, the pane offering "Use slides 2 to 3" |
| `shots/2-attach-your-rows.png` | the four columns pasted, "3 rows attached" |
| `shots/3-preview-one-row.png` | the Nordwind Retail slide merged, beside the card that offers to take it back out |
| `shots/4-before-the-merge.png` | "6 slides will be added after slide 3, leaving 9 slides in the deck", above the button that does it |
| `shots/5-after-the-merge.png` | nine slides in the rail, and the card naming the six it can remove |

`shots/spare-map-the-fields.png` is step 3, the field chips. It is a good
picture and it lost its place to the one after the merge, which is the only shot
that shows the result rather than the process. The name does not start with a
digit, which is how the guard tells an upload from a spare.

### Capturing them again

    node test-kit/driver/reset.mjs
    node test-kit/driver/listing-shots.mjs

against a signed-in PowerPoint for the web with the demo deck open and the
add-in sideloaded. The second walks all five pane steps and writes all five
files. The first is not optional on a deck that has been merged before: the pane
resumes where it was left, so an un-reset run photographs step 1 already saying
"repeat together, 3 times" and offering a "What this run did" disclosure, which
is a state no new customer can be in. The conditions it enforces, each of which cost
a session to learn:

- **1366×768 exactly.** The viewport is emulated before the shutter, not cropped
  after, or the pane's text softens. `Emulation.setDeviceMetricsOverride` and
  the screenshot must share one CDP session, and `scale: "css"` is required or
  the desktop's 1.5 display scale silently produces a 2049-wide file.
- **No avatar, no developer add-in.** `test-kit/driver/shot.mjs` hides the
  account button — which carries a real person's face and their name in its
  aria-label — and any Script Lab tab, and throws if it finds neither, because
  a capture that quietly failed to hide them looks exactly like one that did.
- **A folder with a plain name.** The file name is in the title bar of every
  shot.

Microsoft's guidance is worth reading before the session rather than after:
images should "show real content rather than an empty document", stay legible,
and carry captions that describe the value.

## Logo

There are **three** images, not two, and this section used to say the logo was
"met already, and no work needed". That was true of the two the manifest names
and wrong about the third, which is a separate upload on the listing page and is
marked required.

| Image | Size | Where it goes | Built by |
| --- | --- | --- | --- |
| `IconUrl` | 32×32 | `manifest-prod.xml`, served from the site | `npm run icons` |
| `HighResolutionIconUrl` | 64×64 | `manifest-prod.xml`, served from the site | `npm run icons` |
| Marketplace icon | 300×300, max 512 KB | uploaded to Partner Center by hand | `npm run icons` |

Upload `docs/listing/marketplace-icon-300.png`. It sits beside the screenshots
rather than in `public/assets` because no manifest points at it, so serving it
would put a file on the web that nothing ever requests.

It is the same mark as the ribbon icons, drawn from the same `markPixel`, which
is what Microsoft asks for: "both images should be of the same logo or icon.
This way, the user sees the same logo in Microsoft Marketplace and when the
solution is displayed in Office."

**It is supersampled and the small ones are not.** `markPixel` decides whether a
pixel is inside the rounded square with a hard yes or no. At sixteen pixels the
staircase on the corners is invisible; at three hundred, on a listing page, it
is the first thing you see. The marketplace icon is drawn at 1200 and averaged
down sixteen samples to a pixel. The averaging is premultiplied by alpha, which
is not a detail: averaging the colours raw leaves every corner pixel at rgb
`0,9,19` instead of the true navy `0,37,76`, and that composites over a light
page as a grey fringe around the icon.

`listing.test.ts` pins all of it: 300×300, a whole PNG rather than a truncated
one, under 512 KB, byte-identical to what the build draws, and the same mark at
300 as at 32.

## Notes for certification

Required, and the field says so twice: "Failure to do so results in an automatic
rejection." It is read by a person who has never seen this add-in and has to
make it do something within a few minutes.

```
No test account, license key, or purchase is required. SSF Merge is free. It has
no paid tier, no in-app purchase, no sign-in, and no user accounts. There is
nothing to activate.

After installation the add-in appears on the Home tab, in a group called
SSF Merge. Click "Mail merge" to open the task pane. The pane walks five steps
and states what it will do before it does it.

To test in about two minutes:

1. Create a presentation. Put a title on slide 1. Add slide 2 with the title
   {{Account}} - {{Region}} and a text box reading:
       Renewal {{Renewal|date:d MMM yyyy}}
       Annual value {{Revenue|number:0}} EUR
   Add slide 3 with the title: Next steps for {{Account}}
2. Open Home > Mail merge.
3. Step 1 asks which slides repeat. Enter First slide 2, Last slide 3, continue.
4. Step 2 asks for data. Paste the rows below into the box.
5. Step 3 lists the placeholders found; continue. Step 4 merges one row as a
   preview. Step 5 states how many slides will be added; press the button.

Test data for step 2:

Account,Region,Revenue,Renewal
Nordwind Retail,Nordics,1250000,2026-03-01
Brightline Group,Benelux,880000,2026-04-15
Alpenhof AG,DACH,1640000,2026-05-30

Expected result: six slides are added after slide 3, one pair per row, each
placeholder replaced with that row's value and the original formatting kept.
The pane then offers to remove the slides it just added, which undoes the merge.

A ready-made test deck, the one in the screenshots, can be downloaded from:
https://github.com/dannbleeker/SSF-Merge/raw/main/docs/listing/demo/Quarterly-business-review.pptx
Its template block is slides 2 to 3 and it expects the rows above.

For review: the add-in makes no network requests while it runs. It has no
backend, no telemetry, and no analytics. It reads the open presentation and the
pasted text through the Office JavaScript API, and writes the merged slides back
into the same presentation. The only network traffic is the initial load of the
task pane's own static files from https://ssf-merge.struktureretsundfornuft.dk,
which is served from GitHub Pages.

The add-in requires PowerPointApi 1.2. This is checked at runtime rather than
declared in the manifest, so on a host that lacks it the pane opens and says so
instead of the button silently disappearing.
```

### Why it says what it says

**The credentials sentence goes first** because that is the sentence the field's
own warning is about, and a reviewer who has to hunt for it has already been
given a reason to doubt the rest.

**The rows are comma-separated, not tab-separated.** The pane sniffs the
delimiter and takes tabs, commas or semicolons, so the deck's own `demo/rows.txt`
uses tabs, which is what a spreadsheet paste produces. A cert note is retyped out
of a web form, where a tab is as likely to move the focus as to reach the
clipboard. `test/listing.test.ts` asserts these rows carry the same values as
`demo/rows.txt`, because the failure being guarded against is somebody editing
the demo data and leaving the reviewer pasting rows the deck no longer expects.

**It builds a deck rather than only linking one.** The download is offered second
and as a shortcut. A reviewer who does not want to fetch a binary from GitHub can
still complete the test, and one who takes the link gets the deck the screenshots
were shot from.

**The privacy paragraph is a claim about the source, not marketing.** There is no
`fetch`, `XMLHttpRequest`, `WebSocket` or `sendBeacon` anywhere in `src/`. The
one piece of network traffic that does exist is named rather than glossed over,
because a reviewer watching the network tab will see the task pane's own assets
load and a note claiming "no network requests" flatly would look false.

**Nothing is claimed about which platforms have been tested.** Cross-platform
testing is still open. Saying "tested on PowerPoint for the web" would invite the
question, and saying more than that would not be true.
