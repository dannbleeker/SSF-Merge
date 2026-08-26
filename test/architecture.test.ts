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
  return readFileSync(file, "utf8")
    .replace(/\/\*[\s\S]*?\*\//g, "")
    .split("\n")
    .filter((line) => !/^\s*(\/\/|\*)/.test(line))
    .join("\n");
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

  it("looks values up without walking the prototype chain", () => {
    // A field called __proto__ or constructor is a legal spreadsheet header and
    // arrives from a file the user pasted.
    const src = readFileSync("src/core/data/recordset.ts", "utf8");
    expect(src).toContain("Object.create(null)");
  });
});
