import { describe, expect, it } from "vitest";
import { API_FLOOR, checkFloor, chooseDeckSource, templateOffset } from "../src/host/capability.js";

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
