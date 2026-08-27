/**
 * What this host can do, decided away from the host.
 *
 * Same split as the probe: the question "is this version supported" is answered
 * by Office.js, and everything that follows from the answer is an ordinary
 * function over strings that the suite can check. A capability check written
 * inline in a `PowerPoint.run` callback is a check nobody can test and everybody
 * has to trust.
 */

/** Whether the host supports a given PowerPointApi version. */
export type Supports = (version: string) => boolean;

/**
 * The lowest PowerPointApi version this add-in can work on.
 *
 * Read off the calls it actually makes, against the office-js typings, rather
 * than picked:
 *
 * | call                                     | set |
 * | ---------------------------------------- | --- |
 * | `presentation.slides`                    | 1.2 |
 * | `slides.getCount` / `getItemAt` / `getItem` | 1.2 |
 * | `slide.id`, `slide.delete`               | 1.2 |
 * | `presentation.insertSlidesFromBase64`    | 1.2 |
 * | `slide.tags`                             | 1.3 |
 *
 * So the floor is the tag collection, at **1.3**. `getFileAsync` is a Common
 * API and is not gated by PowerPointApi at all, which is what makes the deck
 * readable on every host that clears this bar.
 *
 * `docs/BACKLOG.md` said 1.4 until this was worked out. Nothing the add-in does
 * needs 1.4, and declaring a floor higher than the truth excludes hosts that
 * would have run it perfectly well.
 */
export const API_FLOOR = "1.3";

export interface Readiness {
  ok: boolean;
  detail: string;
}

/**
 * Whether this host clears the floor.
 *
 * Checked at RUNTIME and never declared in the manifest. A declared requirement
 * set that the host does not meet makes the add-in vanish from the ribbon with
 * no diagnostic at all, so the user sees nothing and has nothing to report; a
 * runtime check can say which version is missing and what that costs them.
 */
export function checkFloor(supports: Supports): Readiness {
  if (supports(API_FLOOR)) {
    return { ok: true, detail: `this host supports PowerPointApi ${API_FLOOR}` };
  }
  return {
    ok: false,
    detail: `SSF Merge needs PowerPointApi ${API_FLOOR} and this host does not have it. Merge metadata lives in slide tags, which arrived in ${API_FLOOR}; without them a merged deck cannot be recognised, undone or re-run.`,
  };
}

/**
 * Where the template's bytes come from.
 *
 * `file` is the floor and works everywhere, including PowerPoint on the web,
 * where the probe read a whole deck back through it. It returns the WHOLE
 * presentation, so the template block sits at its own index inside the package.
 *
 * `subset` is `SlideCollection.exportAsBase64Presentation(ids)` (1.10), which
 * returns a package holding only the slides asked for — masters, layouts and
 * theme included. On a 200-slide deck with a three-slide template that is three
 * slides across the wire instead of two hundred.
 *
 * There is no whole-presentation export to prefer over either. This comment
 * claimed one until the typechecker said otherwise: the 1.10 method exists on
 * `SlideCollection` taking ids, and on the scoped collection `getSelectedSlides`
 * returns — never on `Presentation`. Read the typings, not the shape the API
 * ought to have.
 *
 * Deliberately NOT offering the 1.8 per-slide route. `slide.exportAsBase64`
 * returns one presentation per slide, so a template block of N slides comes
 * back as N packages that would then have to be stitched — masters, layouts,
 * theme and media each appearing N times. That is a real feature and a real
 * amount of code, and it buys nothing the other two do not already give.
 */
export type DeckSource = "subset" | "file";

export interface SourceChoice {
  source: DeckSource;
  detail: string;
}

export function chooseDeckSource(supports: Supports): SourceChoice {
  if (supports("1.10")) {
    return { source: "subset", detail: "exportAsBase64Presentation on just the template slides (PowerPointApi 1.10)" };
  }
  return { source: "file", detail: "getFileAsync, which every host has, returning the whole deck" };
}

/**
 * Where the template block starts inside the package that came back.
 *
 * The two routes disagree, and a caller that assumes either one is wrong half
 * the time: a subset export holds ONLY the template slides, so the block starts
 * at zero, while a whole-deck read holds everything and the block is still
 * wherever it was in the deck. Getting this backwards merges the wrong slides
 * and produces output that looks deliberate.
 */
export function templateOffset(source: DeckSource, blockStartInDeck: number): number {
  return source === "subset" ? 0 : blockStartInDeck;
}

export type BlockIds = { ok: true; ids: string[] } | { ok: false; why: string };

/**
 * The host's OWN ids for a block of slides, out of every id the deck listed.
 *
 * `from` and `to` are 1-based slide numbers, the numbering the thumbnail rail
 * shows and the only one a user can see. `deckIds` is what
 * `slides.load("items/id")` answered, in deck order.
 *
 * This exists because the first version of the merge run did not ask the host
 * at all. It built the ids by counting — `for (let n = from; n <= to; n++)
 * ids.push(String(n))` — and handed `["4", "5", "6"]` to
 * `exportAsBase64Presentation`, whose typings say it *"throws an InvalidArgument
 * exception if provided slide IDs or Slide objects are not found in this
 * collection"*. A slide id on this host looks like `256#3561048925`; `"4"` is
 * not one and never was. `tsc` could not see it, because both sides are
 * `string`. All three answer sheets under `docs/host-answers/` report
 * PowerPointApi up to 1.10, so `chooseDeckSource` returns `subset` on the
 * owner's own host and the first press of the merge button would have thrown.
 *
 * The guard against it coming back is the TYPE: `readTemplate` takes slide
 * numbers now and no caller can pass ids at all. This function is the other
 * half — the host is asked, and what it answers is checked rather than assumed.
 */
export function blockIds(deckIds: string[], from: number, to: number): BlockIds {
  if (!Number.isInteger(from) || !Number.isInteger(to) || from < 1 || to < from) {
    return { ok: false, why: `The template block has to be whole slide numbers, and slide ${from} to ${to} is not.` };
  }
  if (to > deckIds.length) {
    return {
      ok: false,
      why: `The template block is slides ${from} to ${to}, and PowerPoint listed ${deckIds.length} slide(s).`,
    };
  }
  const ids = deckIds.slice(from - 1, to);
  // A blank id is the host refusing to name a slide, which it is documented to
  // do. Passed on it would export the wrong slides or throw somewhere less
  // legible; caught here it is a sentence the pane can show.
  if (ids.some((id) => !id)) {
    return { ok: false, why: `PowerPoint would not name every slide between ${from} and ${to}.` };
  }
  return { ok: true, ids };
}

export type SelectedBlock = { ok: true; from: number; to: number } | { ok: false; why: string };

/**
 * The deck's own id for a slide named by a SELECTION.
 *
 * office-js#2474: a `SlideRange`'s `id` is not roundtrippable — it lacks the
 * `#XYZ` suffix the same slide carries when read from `presentation.slides`, so
 * `slides.getItem(rangeId)` answers InvalidArgument where the deck's own id
 * works. Reported on Windows desktop and closed `not planned`.
 *
 * On the web host this project targets the ids happen to round-trip today, and
 * a sibling add-in has 174 consecutive rounds of a selection read succeeding to
 * say so — which is exactly why this must not be left to luck on a host nobody
 * here has run. The repair is the issue's own observation as a rule: the deck's
 * id is the range's id plus a `#suffix`, so an id absent from the deck's list is
 * matched by that prefix.
 *
 * Exactly one match, or nothing. Two slides answering to one prefix means the
 * assumption behind the repair is wrong on that host, and guessing between them
 * would name the wrong slide — worse than refusing.
 */
export function deckIdForSelectedSlide(rangeId: string, deckIds: string[]): string | undefined {
  if (deckIds.includes(rangeId)) return rangeId;
  const withSuffix = deckIds.filter((id) => id.startsWith(`${rangeId}#`));
  return withSuffix.length === 1 ? withSuffix[0] : undefined;
}

/**
 * What the user has selected, as a template block in SLIDE NUMBERS.
 *
 * The pane speaks in the numbering the thumbnail rail shows, and a selection
 * speaks in ids — so this is the translation, and it is pure because every
 * refusal in it is a sentence the pane shows as it stands.
 *
 * A template block must be CONTIGUOUS: the whole product is "these slides
 * repeat together, in this order, once per row". A selection with a gap in it
 * is refused rather than closed up, because closing it up would silently add
 * slides the user did not pick.
 */
export function blockFromSelection(rangeIds: string[], deckIds: string[]): SelectedBlock {
  if (rangeIds.length === 0) return { ok: false, why: "No slides are selected." };

  const numbers: number[] = [];
  for (const rangeId of rangeIds) {
    const deckId = deckIdForSelectedSlide(rangeId, deckIds);
    // The host named a slide the deck's own list does not carry. Reported
    // rather than skipped: quietly dropping it would build a block out of the
    // slides that happened to resolve.
    if (deckId === undefined) return { ok: false, why: "PowerPoint would not say which slides those are." };
    numbers.push(deckIds.indexOf(deckId) + 1);
  }

  numbers.sort((a, b) => a - b);
  const from = numbers[0] as number;
  const to = numbers[numbers.length - 1] as number;
  if (to - from + 1 !== numbers.length) {
    return {
      ok: false,
      why: `Slides ${from} to ${to} are not all selected — a template block has to be slides that sit next to each other.`,
    };
  }
  return { ok: true, from, to };
}
