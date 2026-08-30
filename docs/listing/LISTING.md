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

| File | Size | Shows |
| --- | --- | --- |
| `01-attach-your-rows.png` | 1366×768, 93 KB | The pane with three rows attached and pictures matched, beside a merged slide |
| `02-see-what-it-will-add.png` | 1366×768, 90 KB | The merge step: what it will add and where, above the button that does it |

**These are placeholders, and should be replaced before submitting.** They are
captured from `test-kit/`'s deck, which exists to TEST the engine rather than to
show it off: the photograph in them is the crop fixture — a blue rectangle with
four yellow dots, positioned to prove a cover-crop keeps the right axis — and
the chart's numbers are `18` and `42` for the same reason. A prospective
customer reads that as a broken image.

Microsoft's guidance is explicit that images should "show real content rather
than an empty document" and that captions describing the value are worth adding.
What the listing wants is a small, honest demo deck — a two-slide template that
looks like something somebody would actually send — merged for three plausible
recipients. That is half an hour of deck-building, and it is the difference
between a listing that demonstrates the product and one that demonstrates the
test suite.

Capture at exactly 1366×768 rather than scaling a larger shot: emulate the size
in the browser first, or the pane's text softens.

## Logo

Met already, and no work needed. AppSource wants **32×32** for the store icon and
**64×64** for its high-resolution partner on a task pane add-in; the manifest now
points `IconUrl` at `icon-32.png` and `HighResolutionIconUrl` at `icon-64.png`,
both of which the icon build already produces.
