import { describe, expect, it } from "vitest";

// @ts-expect-error — a build script, outside `tsconfig.include` and untyped.
import { inComment } from "../scripts/mutate-core.mjs";

/**
 * The mutation runner's own measurement, held to the same standard as the code
 * it measures.
 *
 * A mutant written into a COMMENT changes nothing, so the suite stays green and
 * the run reports a SURVIVOR — a gap in the tests that is not there. This
 * script has now produced two such defects (a whole-line rule that missed a
 * trailing comment, and a copy that left `.git` out and caught everything), and
 * both were invisible in the output: a survivor list reads the same either way.
 */
describe("what the mutation runner will not mutate", () => {
  const at = (source: string, needle: string): number => source.indexOf(needle);

  it("skips a hit inside a comment that TRAILS code on the same line", () => {
    // The shape the first version missed. It rebuilt the whole line and asked
    // whether the LINE began with a comment marker, so a `??` inside a trailing
    // `//` was mutated, changed nothing, and was reported as a gap.
    const src = `const n = a ?? b; // the ?? here is deliberate\n`;
    const trailing = src.lastIndexOf("??");
    expect(inComment(src, trailing), "a ?? inside a trailing comment").toBe(true);
    expect(inComment(src, at(src, "??")), "and the real one beside it is still mutable").toBe(false);
  });

  it("skips a whole-line comment, in every spelling", () => {
    for (const line of ["// a ?? b", " * a ?? b", "/* a ?? b */", "   // a ?? b"]) {
      expect(inComment(`${line}\n`, line.indexOf("??")), line).toBe(true);
    }
  });

  it("does not skip ordinary code", () => {
    const src = `const n = a ?? b;\nconst m = c ?? d;\n`;
    expect(inComment(src, at(src, "??"))).toBe(false);
    expect(inComment(src, src.lastIndexOf("??"))).toBe(false);
  });
});
