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

Maximum 100, recommended 70. This one is 70.

```
One set of slides per row of your table, with fonts and layout intact.
```

It leads with the outcome rather than the name, which is what the guidance asks
for — the name is already on the tile beside it.

## Description

Roughly 380 words. It answers the four questions Microsoft says a description
should answer: how it benefits the user, what is special about it, the different
ways it could be used, and who would use it.

---

Mail merge, for slides instead of letters.

Mark the slides that repeat, paste a table, and get one set of slides per row —
each with that row's name, numbers, dates and picture in place. What took an
afternoon of copy, paste and retype takes a minute, and nothing is retyped, so
nothing is mistyped.

**Your formatting survives, because nothing re-authors it.**

Most tools that fill in a slide rebuild the text, and your fonts, spacing and
theme colours quietly change. This one copies the slides you designed and
replaces only the placeholder, leaving every other property exactly as
PowerPoint wrote it. A deck merged this way looks like the deck you built.

**What it fills in**

- Text anywhere on the slide, including speaker notes
- Numbers and dates, formatted the way you ask: `{{Revenue|number:0}}`,
  `{{Renewal|date:d MMM yyyy}}`
- Pictures, one per row, cropped to fill the frame you drew without distortion
- Chart data and titles, including the workbook behind a chart, so Edit Data
  shows that row's figures
- SmartArt, in both halves it is stored in, so the diagram on screen says the
  right thing

**Before it does anything**

It tells you what it is about to do — how many slides it will add and where
they will land — and shows you one row merged, so you can look at a real result
before committing. Afterwards it will take those slides back out again, and it
names exactly which ones it will remove.

**Who it is for**

Anyone who makes the same deck repeatedly with different names on it: account
reviews, candidate packs, store or regional reports, class or cohort summaries,
certificates, starter packs. If you have ever duplicated a slide twelve times
and edited each copy by hand, this replaces that afternoon.

**Where your data goes**

Nowhere. There is no server, no account and no sign-in, and the add-in makes no
network requests while it runs. Your slides and your table are read inside the
task pane, on your own device, and the merged slides are written straight back
into the presentation you already have open.

---

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
