import { readdirSync, readFileSync, statSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { slideRange } from "../src/core/phrase.js";

/**
 * Comments out, strings KEPT.
 *
 * `withoutTsProse` was the obvious tool and it is the wrong one here: it strips
 * string and template contents as well as comments, so the phrase this guard
 * hunts for disappeared and the guard passed over six files that were breaking
 * it. It was written, it went green, and it was measuring nothing — caught only
 * by putting the offending line back and watching it stay green.
 */
function withoutComments(source: string): string {
  return source.replace(/\/\*[\s\S]*?\*\//g, "").replace(/(^|[^:])\/\/.*$/gm, "$1");
}

function filesUnder(dir: string): string[] {
  return readdirSync(dir).flatMap((entry) => {
    const path = join(dir, entry);
    return statSync(path).isDirectory() ? filesUnder(path) : path.endsWith(".ts") ? [path] : [];
  });
}

describe("saying a slide range", () => {
  it("says slide, singular, when the range is one slide", () => {
    expect(slideRange(3, 3)).toBe("slide 3");
    expect(slideRange(1, 1)).toBe("slide 1");
  });

  it("says slides, with both ends, when it is more than one", () => {
    expect(slideRange(2, 5)).toBe("slides 2 to 5");
  });

  it("never produces a range with itself at both ends", () => {
    // The defect this function exists for. "slides 2 to 2" reads as something
    // the user mistyped rather than a sentence the product meant, and it is the
    // commonest block there is: one slide per row.
    for (let n = 1; n <= 40; n++) expect(slideRange(n, n)).not.toContain(" to ");
  });
});

describe("nothing builds a slide range by hand", () => {
  /**
   * The guard, and the reason the function above is in `core` rather than in
   * the pane where it was first needed.
   *
   * Seven files spelled this phrase out themselves. Six agreed with each other
   * and were wrong for a one-slide block; the seventh, `blockName`, was right,
   * which is how the manual ended up with a screenshot whose heading said
   * "Slide 3" directly above a button saying "Use slides 3 to 3".
   *
   * Fixing six copies without this guard just resets the clock. The eighth is
   * written by somebody who never saw any of them.
   */
  const SPELLED_OUT = /slides? \$\{[^}]*\} to \$\{/;

  it("is true of every source file except the one that defines it", () => {
    const offenders = filesUnder("src")
      // Comments quote the old wording on purpose — several explain the very
      // defect this forbids — so the prose is stripped before looking.
      .filter((f) => SPELLED_OUT.test(withoutComments(readFileSync(f, "utf8"))))
      .filter((f) => !f.endsWith(join("core", "phrase.ts")));
    expect(offenders, "these build the phrase themselves; call slideRange").toEqual([]);
  });
});
