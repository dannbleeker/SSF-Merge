import { execFileSync } from "node:child_process";
import { mkdtempSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { describe, expect, it } from "vitest";

// @ts-expect-error — a plain .mjs tool with no types. Narrowed here rather than
// at every call, so an arrow returning it is not an unchecked `any`.
import { isMain as untyped } from "../scripts/is-main.mjs";

const isMain = untyped as (metaUrl: string, argv?: string[]) => boolean;

/**
 * The predicate every tool CLI in this repo guards its `main()` with.
 *
 * It had no test at all, which is the worst place for that: when this answers
 * false the script prints nothing and exits 0, and a clean exit reads exactly
 * like a pass. A sibling project lost three tool CLIs to that for months
 * without noticing, and this module's own docstring is about it.
 */
describe("whether a module is the entry point", () => {
  it("says yes for the file that was run", () => {
    const file = resolve("scripts/check-release.mjs");
    expect(isMain(pathToFileURL(file).href, ["node", file])).toBe(true);
  });

  it("says no for a module that was merely imported", () => {
    expect(
      isMain(pathToFileURL(resolve("scripts/is-main.mjs")).href, ["node", resolve("scripts/test-count.mjs")]),
    ).toBe(false);
  });

  it("says no when there is no entry at all", () => {
    // `node -e` and a REPL both leave `argv[1]` undefined, and reading `.href`
    // of that would throw inside a guard whose whole job is to be quiet.
    expect(isMain("file:///anything.mjs", ["node"])).toBe(false);
  });

  it("still says yes through a SYMLINK to the file", () => {
    /**
     * Node resolves `import.meta.url` to the link's TARGET and leaves `argv[1]`
     * as the link, so the two never matched: the CLI did nothing and exited 0.
     * `npm run` passes a direct path and never meets it; `npm link`, a
     * `node_modules/.bin` shim and a global install all go through one.
     *
     * Driven through a real `node` rather than by calling the predicate,
     * because the thing under test is what Node does with the two values — a
     * unit call would be asserting my own model of it.
     */
    const dir = mkdtempSync(join(tmpdir(), "ssf-is-main-"));
    const real = join(dir, "real.mjs");
    const link = join(dir, "link.mjs");
    writeFileSync(
      real,
      `import { isMain } from ${JSON.stringify(resolve("scripts/is-main.mjs"))};\n` +
        `console.log(isMain(import.meta.url) ? "MAIN" : "NOT MAIN");\n`,
    );
    symlinkSync(real, link);

    const ran = (entry: string) => execFileSync(process.execPath, [entry], { encoding: "utf8" }).trim();
    expect(ran(real), "directly").toBe("MAIN");
    expect(ran(link), "through a symlink").toBe("MAIN");
  });

  it("answers for a path that is not on disk rather than throwing", () => {
    // `realpathSync` raises for a missing file, and losing the answer there
    // would be the same silence by another route.
    const ghost = resolve("scripts/nothing-here.mjs");
    const ask = () => isMain(pathToFileURL(ghost).href, ["node", ghost]);
    expect(ask).not.toThrow();
    expect(ask()).toBe(true);
  });
});
