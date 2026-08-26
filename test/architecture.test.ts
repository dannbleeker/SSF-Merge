import { readdirSync, readFileSync, statSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";

function filesUnder(dir: string): string[] {
  return readdirSync(dir).flatMap((entry) => {
    const path = join(dir, entry);
    return statSync(path).isDirectory() ? filesUnder(path) : path.endsWith(".ts") ? [path] : [];
  });
}

/**
 * The file with its prose removed.
 *
 * The first version of this guard matched the words "Office.js" and
 * "PowerPoint.run" in the comments that explain WHY the engine avoids them, so
 * it failed on four files that were entirely correct. A guard that goes red for
 * the wrong reason teaches the next reader to widen it until it goes green.
 */
function codeOf(file: string): string {
  return (
    readFileSync(file, "utf8")
      .replace(/\/\*[\s\S]*?\*\//g, "")
      .split("\n")
      .filter((line) => !/^\s*(\/\/|\*)/.test(line))
      .join("\n")
      // String literals go too, for the same reason the comments did. A verdict
      // that names office-js#6105 in its text is a sentence about an issue, not
      // a dependency on it, and the first version of the src/host guard failed
      // on exactly that — a file with no imports at all.
      .replace(/`(?:[^`\\]|\\.)*`/g, '""')
      .replace(/"(?:[^"\\\n]|\\.)*"/g, '""')
      .replace(/'(?:[^'\\\n]|\\.)*'/g, '""')
  );
}

describe("src/core", () => {
  it("imports nothing from Office.js", () => {
    // The seam the whole design rests on. The engine takes bytes and records
    // and returns bytes, which is what lets it run in the pane, in a CLI and in
    // this suite with no PowerPoint anywhere. One import here and the merge can
    // only ever be tested inside a host that is documented to lie about ids.
    const offenders = filesUnder("src/core").filter((f) => codeOf(f).match(/\bOffice\.|\bPowerPoint\.|office-js/));
    expect(offenders).toEqual([]);
  });

  it("keeps Office.js out of src/host too, so every decision stays testable", () => {
    // src/host is the DECISIONS: which version floor the host clears, where the
    // template's bytes come from, whether an insert did what it said, which
    // slides an undo may remove. src/office is the calls. The whole reason the
    // probe could answer anything is that its judgements lived outside the
    // PowerPoint.run callback and could be checked in CI; one Office.js import
    // here and the same rules become untestable again.
    const offenders = filesUnder("src/host").filter((f) => codeOf(f).match(/\bOffice\.|\bPowerPoint\.|office-js/));
    expect(offenders).toEqual([]);
  });

  it("keeps the decisions out of src/office, so none of them hide in a callback", () => {
    // The other direction, and the one that rots quietly: a rule reimplemented
    // inline next to the call it guards looks tidier and is untestable. Every
    // judgement src/office needs is imported from src/host — the version floor,
    // the source choice, the insert reading, the sweep plan — so a file here
    // with no such import is either trivial or has started deciding for itself.
    const office = filesUnder("src/office");
    expect(office.length).toBeGreaterThan(0);
    for (const file of office) {
      // The RAW source, not codeOf: an import specifier IS a string literal,
      // and codeOf strips those so the Office.js check above cannot trip over
      // prose. Two questions, two readings of the same file.
      expect(readFileSync(file, "utf8"), `${file} decides nothing on its own`).toMatch(/from "\.\.\/host\//);
    }
  });

  it("looks values up without walking the prototype chain", () => {
    // A field called __proto__ or constructor is a legal spreadsheet header and
    // arrives from a file the user pasted.
    const src = readFileSync("src/core/data/recordset.ts", "utf8");
    expect(src).toContain("Object.create(null)");
  });
});
