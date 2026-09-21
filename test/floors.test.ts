/**
 * The floors this project declares, held against what it actually installs.
 *
 * Every other dependency check here reads the INSTALLED tree: `npm audit`
 * reports on what the lockfile resolved, the suite runs against what `npm ci`
 * put on disk, and `docs/DEPENDENCY-ALERTS.md` records a reading of whatever
 * Dependabot proposed. None of them reads the lower bound of a caret, and that
 * is where `@xmldom/xmldom` sat at `^0.9.8` — seventeen advisories, one of them
 * high severity, against the parser that reads every part of a user-supplied
 * deck — while the lockfile held 0.9.12 and every gate on this repository was
 * green. It had been that way since the day the dependency was added.
 *
 * The rule is in `scripts/dependency-floors.mjs` so there is one copy of it.
 * What it can and cannot see is written down there; the short version is that
 * it checks ALIGNMENT, not advisories, because the advisory database changes
 * daily and a test built on it would go red on a day nobody touched the code.
 * Aligned, the floor is the version the whole gate has already run.
 *
 * Five of the tests below feed it inputs that must be REFUSED. That is
 * deliberate and it is the lesson of the four gates this repository found could
 * not fail: a scan over the real files returning "no problems" is also exactly
 * what a scanner that has stopped matching returns, so the real-file test
 * checks what it reached before believing what it proved.
 */
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
// @ts-expect-error — a plain .mjs tool with no types. The rules live THERE so
// the suite and any tool asking the same question cannot read different ones.
import { floorOf, floorProblems } from "../scripts/dependency-floors.mjs";

type Manifest = {
  dependencies?: Record<string, string>;
  devDependencies?: Record<string, string>;
};
type Lock = {
  packages?: Record<
    string,
    { version?: string; dependencies?: Record<string, string>; devDependencies?: Record<string, string> }
  >;
};
type Report = { problems: string[]; scanned: number };

const read = (path: string) => JSON.parse(readFileSync(path, "utf8")) as unknown;
const MANIFEST = read("package.json") as Manifest;
const LOCK = read("package-lock.json") as Lock;

const check = (manifest: Manifest, lock: Lock): Report => floorProblems({ manifest, lock }) as Report;

describe("floorOf", () => {
  it("reads the lowest version a caret, a tilde and a bare version admit", () => {
    // The anchor: three answers known from outside the code under test.
    expect(floorOf("^1.2.3")).toBe("1.2.3");
    expect(floorOf("~0.9.12")).toBe("0.9.12");
    expect(floorOf("5.0.1")).toBe("5.0.1");
    expect(floorOf("^2.0.0-beta.4")).toBe("2.0.0-beta.4");
  });

  it("answers null for a range whose floor it cannot read, rather than guessing one", () => {
    for (const range of ["*", ">=1.0.0", "1.x", "^1.2", "1.2.3 || 2.0.0", "latest", "file:../x", "workspace:*", ""]) {
      expect(floorOf(range), range).toBeNull();
    }
    expect(floorOf(undefined)).toBeNull();
  });
});

describe("the declared floors", () => {
  it("are the versions the lockfile installs", () => {
    const { problems } = check(MANIFEST, LOCK);
    expect(problems).toEqual([]);
  });

  /**
   * The vacuity half, and it is not decoration.
   *
   * `problems: []` is the answer for a clean repository AND the answer for a
   * scanner reading the wrong key, a manifest that parsed to `{}`, or a rename
   * upstream. The docProps guard in this suite was open in exactly this
   * dimension for weeks: it asserted an empty list and an empty list was also
   * what a matcher that had stopped matching returned.
   */
  it("were actually read — the scan reaches every direct dependency by name", () => {
    const { scanned } = check(MANIFEST, LOCK);
    const declared = {
      ...(MANIFEST.dependencies ?? {}),
      ...(MANIFEST.devDependencies ?? {}),
    };
    const names = Object.keys(declared);

    expect(scanned).toBe(names.length);
    expect(scanned).toBeGreaterThan(10);
    // The two runtime dependencies by name. They are the ones that reach a
    // user's bytes, and a scan that has stopped seeing them is the failure
    // this test exists to make loud.
    expect(names).toContain("@xmldom/xmldom");
    expect(names).toContain("jszip");
  });
});

describe("what the check refuses", () => {
  it("reports a floor below the installed version, which is the defect it was written for", () => {
    // The shape found on 2026-09-21, reduced: a caret four releases behind the
    // version the same install resolved.
    const { problems } = check(
      { dependencies: { "@xmldom/xmldom": "^0.9.8" } },
      {
        packages: {
          "": { dependencies: { "@xmldom/xmldom": "^0.9.8" } },
          "node_modules/@xmldom/xmldom": { version: "0.9.12" },
        },
      },
    );

    expect(problems).toHaveLength(1);
    expect(problems[0]).toContain("@xmldom/xmldom");
    expect(problems[0]).toContain("0.9.8");
    expect(problems[0]).toContain("0.9.12");
  });

  it("reports a floor ABOVE the installed version too, which is the direction a bound written to keep something out forgets", () => {
    const { problems } = check(
      { devDependencies: { vitest: "^5.0.1" } },
      { packages: { "": { devDependencies: { vitest: "^5.0.1" } }, "node_modules/vitest": { version: "5.0.0" } } },
    );

    expect(problems).toHaveLength(1);
    expect(problems[0]).toContain("vitest");
  });

  it("refuses a range it cannot read a floor from, rather than skipping it quietly", () => {
    const { problems, scanned } = check(
      { dependencies: { somepkg: ">=1.0.0" } },
      { packages: { "": {}, "node_modules/somepkg": { version: "1.4.0" } } },
    );

    expect(scanned).toBe(1);
    expect(problems).toHaveLength(1);
    expect(problems[0]).toContain("cannot read a floor");
  });

  it("reports a dependency the lockfile does not install", () => {
    const { problems } = check({ dependencies: { ghost: "^1.0.0" } }, { packages: { "": {} } });

    expect(problems).toHaveLength(1);
    expect(problems[0]).toContain("no installed version");
  });

  it("reports a manifest range the lockfile's own copy contradicts, which is what a hand edit without npm install leaves", () => {
    const { problems } = check(
      { devDependencies: { prettier: "^3.9.8" } },
      { packages: { "": { devDependencies: { prettier: "^3.9.6" } }, "node_modules/prettier": { version: "3.9.8" } } },
    );

    expect(problems).toHaveLength(1);
    expect(problems[0]).toContain("lockfile's own");
  });

  it("says nothing about a manifest with no dependencies at all, and says it scanned nothing", () => {
    // The honest empty case, kept so the two meanings of "no problems" stay
    // distinguishable in the report rather than only in this file's prose.
    const { problems, scanned } = check({}, { packages: { "": {} } });
    expect(problems).toEqual([]);
    expect(scanned).toBe(0);
  });
});
