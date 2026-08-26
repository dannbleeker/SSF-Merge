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
