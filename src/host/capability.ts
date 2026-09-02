import { slideRange } from "../core/phrase.js";
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
 * | `slides.load("items/id")`                | 1.2 |
 *
 * So the floor is **1.2**. `getFileAsync` is a Common API and is not gated by
 * PowerPointApi at all, which is what makes the deck readable on every host
 * that clears this bar.
 *
 * This said 1.4, then 1.3, and both were wrong in the same direction. 1.3 was
 * justified by `slide.tags`, which at the time nothing called: merge metadata
 * is written into the PACKAGE, as `ppt/tags/tagN.xml` with a relationship from
 * the slide, which is the whole reason the engine never asks the host for
 * anything it can put in the file.
 *
 * `runTagsAt` DOES read `slide.tags` now — the undo asks each slide in its
 * range whether this run made it — and the floor is still 1.2, deliberately.
 * That read is guarded where it is made and answers an empty list below 1.3, so
 * an older host loses the identity check and keeps the merge. Declaring 1.3
 * would exclude a host that runs the add-in perfectly well, and it would do it
 * SILENTLY: `checkFloor` tells the user their PowerPoint is too old when it is
 * not. (This paragraph asserted "a repo-wide grep for `.tags` finds no caller"
 * for as long as the caller existed. Re-run the grep before repeating it.)
 *
 * Three calls sit ABOVE the floor and are each guarded where they are used,
 * never declared: `exportAsBase64Presentation` (1.10, `chooseDeckSource` falls
 * back to the file API), `getSelectedSlides` (1.5, `canSelectSlides` hides the
 * control) and `Slide.tags` (1.3, `runTagsAt` answers empty). An optional call
 * is not a floor — but the third one changes what a button DOES rather than
 * whether it is there, so `docs/MANUAL.md` names it under the undo as well.
 */
export const API_FLOOR = "1.2";

/**
 * Whether this host can say which slides are selected.
 *
 * `getSelectedSlides` is **PowerPointApi 1.5** and the floor is 1.2, so it is
 * an EXTRA rather than a requirement: on an older host the two slide-number
 * boxes still work and the shortcut is simply not offered. A control that
 * always fails is worse than one that is not there.
 *
 * This was shipped unguarded — the call went in on the strength of a sibling
 * project's rounds showing it is not WEDGED, without anyone asking which
 * version introduced it. Being safe to call and being present are different
 * questions.
 */
export function canSelectSlides(supports: Supports): boolean {
  return supports("1.5");
}

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
    detail: `SSF Merge needs PowerPointApi ${API_FLOOR} and this host does not have it. Reading the deck, inserting the merged slides and taking them back again all need it; without it there is nothing the add-in can do.`,
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
    // Not `slideRange`: this echoes what was TYPED rather than naming a range
    // of slides, and "slides 1.5 to 2.5 is not" puts a plural subject on a
    // singular verb. Dropping the word is what makes it read.
    return { ok: false, why: `The template block has to be whole slide numbers, and ${from} to ${to} is not.` };
  }
  if (to > deckIds.length) {
    return {
      ok: false,
      why: `The template block is ${slideRange(from, to)}, and PowerPoint listed ${deckIds.length} slide(s).`,
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
 * a sibling add-in's selection read has succeeded in every archived round to
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
    // ONCE each. The contiguity test below compares `to - from + 1` against how
    // many numbers are in this list, and a count is not an alignment: name one
    // slide twice and the count covers a gap it should have caught.
    //
    // Both directions were wrong, and one of them silently. A selection of
    // slides 1 and 3 with slide 1 named twice counted three numbers spanning
    // three slides and came back "slides 1 to 3" — building the template out of
    // slide 2, which the user never selected, and which is the exact thing the
    // comment above says must not happen. The other direction merely refused a
    // selection that was fine.
    const number = deckIds.indexOf(deckId) + 1;
    if (!numbers.includes(number)) numbers.push(number);
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

/** Every PowerPointApi version this add-in ever asks about. */
export const KNOWN_SETS = ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10"] as const;

export interface Environment {
  build: string;
  platform: string;
  /**
   * WHICH HOST — "PowerPoint", "Word". `Office.context.diagnostics.host`.
   *
   * This was filled from `diagnostics.version` and named `host`, so the one
   * field that answers "which application am I in" carried a build number, and
   * the question this line exists to answer went unasked. The version has its
   * own field below.
   */
  host: string;
  /** The Office build, `Office.context.diagnostics.version`. */
  officeVersion: string;
  /** EVERY set the host publishes, not only the ones we gate on. */
  sets: string[];
  floor: string;
  clearsFloor: boolean;
  deckSource: DeckSource;
  canSelect: boolean;
}

/**
 * What this host IS, as one line at the top of a run.
 *
 * The first three questions of any investigation are which build, which host
 * and which API — and answering them costs one line that a reader gets for
 * free. A sibling project's rounds carry this and it has paid for itself; the
 * rounds that DON'T are the ones where somebody had to ask.
 *
 * **Every set the host publishes, not only the ones we gate on.** The gap
 * between what a host HAS and what this add-in uses is where the next unusable
 * API is hiding, and a list filtered to what we already call can never show it.
 *
 * Pure: the caller passes the readings in, so the whole thing is checkable
 * without a PowerPoint. `src/office` is where they are gathered — and each one
 * individually, because a host that throws reading its own version must not
 * cost the round the rest of the line.
 */
export function environmentLine(o: {
  build?: string;
  platform?: string;
  host?: string;
  officeVersion?: string;
  supports: Supports;
}): Environment {
  return {
    // "unknown" rather than an empty string: a blank in a run log reads as a
    // field nobody wrote, and this one was written and had no answer.
    build: o.build ?? "unknown",
    platform: o.platform ?? "unknown",
    host: o.host ?? "unknown",
    officeVersion: o.officeVersion ?? "unknown",
    sets: KNOWN_SETS.filter((v) => o.supports(v)),
    floor: API_FLOOR,
    clearsFloor: checkFloor(o.supports).ok,
    deckSource: chooseDeckSource(o.supports).source,
    canSelect: canSelectSlides(o.supports),
  };
}
