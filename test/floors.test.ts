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
// @ts-expect-error — the weekly sweep, same reason. Its network half is a
// scheduled job; the refusals and the control arm are tested here.
import { advisoryReport, floorManifest, main as mainSweep } from "../scripts/audit-floors.mjs";

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

/**
 * The weekly sweep, in the half that can be tested at all.
 *
 * `scripts/audit-floors.mjs` asks the registry and the advisory database what
 * the declared floors carry, which is why it is a scheduled job rather than a
 * test. Everything below runs against injected answers instead: the shapes it
 * must REFUSE, and the control arm it will not report without. Those are the
 * parts a network outage would otherwise turn into a clean-looking sweep, and
 * they are exactly what a live run cannot demonstrate on a good day.
 */
describe("the floor audit sweep", () => {
  const CONTROL_PINS = JSON.stringify({ "@xmldom/xmldom": "0.9.8" });
  const clean = { findings: [], counted: { total: 0 }, problems: [] };
  const controlSaw = (n: number) => ({
    findings: [{ name: "@xmldom/xmldom", severity: "high", direct: true, advisories: n, titles: [] }],
    counted: { total: 1 },
    problems: [],
  });

  /** A `main` wired to answers of our choosing, with both streams captured. */
  const runSweep = async (manifest: unknown, answer: (pins: Record<string, string>) => unknown) => {
    let out = "";
    let err = "";
    const code = (await mainSweep([], {
      read: () => JSON.stringify(manifest),
      audit: answer,
      out: { write: (s: string) => (out += s) },
      err: { write: (s: string) => (err += s) },
    })) as number;
    return { code, out, err };
  };

  it("pins every declared range to its own lower bound, and keeps the ones it cannot read", () => {
    const { pinned, unreadable } = floorManifest({
      dependencies: { "@xmldom/xmldom": "^0.9.12" },
      devDependencies: { vitest: "~5.0.1", odd: ">=2" },
    }) as { pinned: Record<string, string>; unreadable: { name: string; range: string }[] };

    expect(pinned).toEqual({ "@xmldom/xmldom": "0.9.12", vitest: "5.0.1" });
    expect(unreadable).toEqual([{ name: "odd", range: ">=2" }]);
  });

  it("refuses a report it cannot read, rather than calling it clean", () => {
    const refusals = [
      advisoryReport(null),
      advisoryReport("not json"),
      advisoryReport({ auditReportVersion: 3, metadata: { vulnerabilities: { total: 0 } } }),
      advisoryReport({ auditReportVersion: 2 }),
    ] as { problems: string[]; findings: unknown[] }[];

    for (const refusal of refusals) {
      expect(refusal.problems.length).toBeGreaterThan(0);
      expect(refusal.findings).toEqual([]);
    }
  });

  it("reads a version 2 report, counting only the advisory objects in `via`", () => {
    // `via` holds advisory objects AND plain strings naming the package a
    // finding came through. Counting the strings would inflate every number
    // this sweep prints.
    const parsed = advisoryReport({
      auditReportVersion: 2,
      metadata: { vulnerabilities: { total: 1 } },
      vulnerabilities: {
        thing: {
          severity: "high",
          isDirect: true,
          via: ["some-other-package", { title: "a real advisory", url: "https://example.invalid/1" }],
        },
      },
    }) as { problems: string[]; findings: { name: string; advisories: number; titles: string[]; direct: boolean }[] };

    const [first] = parsed.findings;
    expect(parsed.problems).toEqual([]);
    expect(parsed.findings).toHaveLength(1);
    expect(first?.advisories).toBe(1);
    expect(first?.titles).toEqual(["a real advisory"]);
    expect(first?.direct).toBe(true);
  });

  it("will not report at all when the control arm sees nothing, which is what an unreachable database looks like", async () => {
    const { code, out, err } = await runSweep({ dependencies: { "@xmldom/xmldom": "^0.9.12" } }, () => clean);

    expect(code).toBe(1);
    // Not one word about the real floors. A sweep that cannot see advisories
    // has no opinion about whether there are any.
    expect(out).toBe("");
    expect(err).toContain("control arm failed");
  });

  it("will not report when the control arm sees FEWER advisories than were measured", async () => {
    const { code, err } = await runSweep({ dependencies: { "@xmldom/xmldom": "^0.9.12" } }, (pins) =>
      JSON.stringify(pins) === CONTROL_PINS ? controlSaw(1) : clean,
    );

    expect(code).toBe(1);
    expect(err).toContain("reported 1 advisories");
  });

  it("reports nothing wrong when the control passes and every floor is clean", async () => {
    const { code, out } = await runSweep({ dependencies: { "@xmldom/xmldom": "^0.9.12" } }, (pins) =>
      JSON.stringify(pins) === CONTROL_PINS ? controlSaw(17) : clean,
    );

    expect(code).toBe(0);
    expect(out).toContain("No advisories against any declared floor");
    expect(out).toContain("The sweep can see advisories");
  });

  it("reports a floor that carries advisories, and exits 3 rather than 0 or 1", async () => {
    // A package OTHER than the control's, deliberately. Pinned at `^0.9.8`
    // the manifest's own floor set is byte-identical to the control's pins,
    // so a stub keyed on them answers the control twice and the assertion
    // about the finding passes on the control's empty title list. The first
    // draft of this test did exactly that.
    const { code, out } = await runSweep({ dependencies: { "some-parser": "^1.0.0" } }, (pins) =>
      JSON.stringify(pins) === CONTROL_PINS
        ? controlSaw(17)
        : {
            findings: [
              { name: "some-parser", severity: "high", direct: true, advisories: 17, titles: ["quadratic memory"] },
            ],
            counted: { total: 1 },
            problems: [],
          },
    );

    expect(code).toBe(3);
    expect(out).toContain("Floors carrying advisories");
    expect(out).toContain("17 advisories");
    expect(out).toContain("quadratic memory");
  });

  it("treats a range it could not pin as a finding, because a package it skipped is one it did not audit", async () => {
    const { code, out } = await runSweep({ dependencies: { odd: "*" } }, (pins) =>
      JSON.stringify(pins) === CONTROL_PINS ? controlSaw(17) : clean,
    );

    expect(code).toBe(3);
    expect(out).toContain("no readable floor");
    expect(out).toContain("was NOT audited");
  });
});
