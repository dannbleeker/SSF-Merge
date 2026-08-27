import { describe, expect, it } from "vitest";
import {
  API_FLOOR,
  blockFromSelection,
  blockIds,
  checkFloor,
  chooseDeckSource,
  deckIdForSelectedSlide,
  templateOffset,
} from "../src/host/capability.js";

/** A host that supports every version up to and including `top`. */
function upTo(top: string) {
  const rank = (v: string) => v.split(".").map(Number);
  return (v: string) => {
    const [a = 0, b = 0] = rank(v);
    const [x = 0, y = 0] = rank(top);
    return a < x || (a === x && b <= y);
  };
}

describe("the version floor", () => {
  it("is 1.3, which is where slide tags arrived", () => {
    // Read off the calls the add-in makes rather than picked. Everything else
    // it needs — slides, getItemAt, insertSlidesFromBase64, slide.delete — is
    // 1.2, and getFileAsync is a Common API that PowerPointApi does not gate.
    expect(API_FLOOR).toBe("1.3");
  });

  it("passes a host that has it", () => {
    expect(checkFloor(upTo("1.3")).ok).toBe(true);
  });

  it("fails a host that stops at 1.2, and says what that costs", () => {
    const r = checkFloor(upTo("1.2"));
    expect(r.ok).toBe(false);
    // The message has to be usable by whoever reads it in the pane: a bare
    // "unsupported" sends them nowhere.
    expect(r.detail).toContain(API_FLOOR);
    expect(r.detail).toMatch(/tags/i);
  });

  it("passes every host above the floor", () => {
    for (const top of ["1.3", "1.4", "1.5", "1.8", "1.10"]) {
      expect(checkFloor(upTo(top)).ok, `host up to ${top}`).toBe(true);
    }
  });
});

describe("where the template's bytes come from", () => {
  it("exports just the template slides when the host has 1.10", () => {
    expect(chooseDeckSource(upTo("1.10")).source).toBe("subset");
  });

  it("falls back to reading the whole file below 1.10", () => {
    for (const top of ["1.3", "1.8", "1.9"]) {
      expect(chooseDeckSource(upTo(top)).source, `host up to ${top}`).toBe("file");
    }
  });

  it("names the call it chose, because a route nobody can see is a route nobody can debug", () => {
    expect(chooseDeckSource(upTo("1.10")).detail).toContain("exportAsBase64Presentation");
    expect(chooseDeckSource(upTo("1.3")).detail).toContain("getFileAsync");
  });
});

describe("where the template block starts in the package that came back", () => {
  it("is zero for a subset export, which holds only the template", () => {
    expect(templateOffset("subset", 17)).toBe(0);
  });

  it("is the deck position for a whole-file read, which holds everything", () => {
    expect(templateOffset("file", 17)).toBe(17);
  });

  it("differs between the routes for any block that is not first", () => {
    // The whole reason this function exists. A caller that assumes either
    // answer merges the wrong slides on half the hosts, and the output looks
    // deliberate rather than broken.
    for (const start of [1, 2, 9, 200]) {
      expect(templateOffset("subset", start)).not.toBe(templateOffset("file", start));
    }
    // And they agree exactly when the block is first, which is the case that
    // would hide the bug in every hand test somebody runs.
    expect(templateOffset("subset", 0)).toBe(templateOffset("file", 0));
  });
});

describe("the host's own ids for a block of slides", () => {
  // A slide id on this host looks like "256#3561048925". The first merge run
  // never asked: it counted, and handed exportAsBase64Presentation
  // ["4", "5", "6"] — strings on both sides, so tsc saw nothing and the typings
  // say it throws InvalidArgument for an id it cannot find. All three answer
  // sheets report PowerPointApi 1.10, so that was the owner's own host on the
  // first press of the merge button.
  const deck = ["256#3561048925", "257#1897035307", "258#2230304510", "259#4123571115"];

  it("takes the ids the host listed, never numbers", () => {
    const chosen = blockIds(deck, 2, 3);
    expect(chosen.ok && chosen.ids).toEqual(["257#1897035307", "258#2230304510"]);
  });

  it("counts BOTH ends of the block", () => {
    // Slides 1 to 4 is four slides. The off-by-one that would export three.
    const chosen = blockIds(deck, 1, 4);
    expect(chosen.ok && chosen.ids).toHaveLength(4);
  });

  it("takes a one-slide block", () => {
    const chosen = blockIds(deck, 3, 3);
    expect(chosen.ok && chosen.ids).toEqual(["258#2230304510"]);
  });

  it("refuses a block that runs past what the host listed, and says both numbers", () => {
    const chosen = blockIds(deck, 3, 9);
    expect(chosen.ok).toBe(false);
    expect(!chosen.ok && chosen.why).toContain("3 to 9");
    expect(!chosen.ok && chosen.why).toContain("4");
  });

  it("refuses a blank id rather than exporting whatever comes back", () => {
    // The host is documented to refuse to name things. Passed on, an empty id
    // either throws somewhere less legible or exports the wrong slides.
    expect(blockIds(["a", "", "c"], 1, 3).ok).toBe(false);
  });

  it("refuses numbers that are not whole slides", () => {
    for (const [from, to] of [
      [0, 2],
      [2, 1],
      [1.5, 3],
      [Number.NaN, 2],
    ] as const) {
      expect(blockIds(deck, from, to).ok, `${from} to ${to}`).toBe(false);
    }
  });
});

describe("the deck's own id for a slide named by a selection", () => {
  // office-js#2474: a SlideRange's id lacks the `#XYZ` suffix the same slide
  // carries when read from `presentation.slides`, so `getItem(rangeId)`
  // answers InvalidArgument where the deck's id works. Closed `not planned`.
  const deck = ["256#3561048925", "257#1897035307", "258#2230304510"];

  it("takes an id that is already the deck's", () => {
    expect(deckIdForSelectedSlide("257#1897035307", deck)).toBe("257#1897035307");
  });

  it("matches a suffix-less range id by its prefix", () => {
    expect(deckIdForSelectedSlide("257", deck)).toBe("257#1897035307");
  });

  it("refuses when two slides answer to one prefix", () => {
    // Guessing between them would name the wrong slide, which is worse than
    // refusing — and it means the assumption behind the whole repair is wrong
    // on that host.
    expect(deckIdForSelectedSlide("25", ["25#a", "25#b"])).toBeUndefined();
  });

  it("refuses an id the deck does not carry at all", () => {
    expect(deckIdForSelectedSlide("999", deck)).toBeUndefined();
  });
});

describe("the template block a selection names", () => {
  const deck = ["256#a", "257#b", "258#c", "259#d", "260#e"];

  it("reads contiguous slides as a block, in the rail's numbering", () => {
    expect(blockFromSelection(["257#b", "258#c", "259#d"], deck)).toEqual({ ok: true, from: 2, to: 4 });
  });

  it("does not care what order the host listed them in", () => {
    expect(blockFromSelection(["259#d", "257#b", "258#c"], deck)).toEqual({ ok: true, from: 2, to: 4 });
  });

  it("reads a single slide as a one-slide block", () => {
    expect(blockFromSelection(["258#c"], deck)).toEqual({ ok: true, from: 3, to: 3 });
  });

  it("works through the suffix-less ids office-js#2474 describes", () => {
    expect(blockFromSelection(["257", "258"], deck)).toEqual({ ok: true, from: 2, to: 3 });
  });

  it("REFUSES a selection with a gap rather than closing it up", () => {
    // The whole product is "these slides repeat together". Closing the gap
    // would silently add a slide the user did not pick, to every row.
    const read = blockFromSelection(["257#b", "259#d"], deck);
    expect(read.ok).toBe(false);
    expect(!read.ok && read.why).toContain("next to each other");
    expect(!read.ok && read.why).toContain("2 to 4");
  });

  it("refuses an empty selection", () => {
    expect(blockFromSelection([], deck)).toEqual({ ok: false, why: "No slides are selected." });
  });

  it("refuses rather than dropping a slide the deck will not name", () => {
    // Quietly skipping it would build a block out of whichever slides happened
    // to resolve — a block the user never selected.
    const read = blockFromSelection(["257#b", "999#z"], deck);
    expect(read.ok).toBe(false);
    expect(!read.ok && read.why).toContain("would not say");
  });
});
