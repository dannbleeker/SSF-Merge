/**
 * Whether the manifest's declared floors match the versions actually installed.
 *
 * A range is a claim about the WORST version this project has said it will
 * accept, and nothing in the gate reads it. `npm audit` audits what RESOLVED —
 * it never looks at the lower bound of a caret — so a floor can name a version
 * with open advisories against it while every check on the repository stays
 * green. That is not hypothetical: on 2026-09-21 `@xmldom/xmldom` was declared
 * `^0.9.8` against a lockfile holding 0.9.12, and 0.9.8 carried 17 advisories,
 * one of them high severity, against the parser that reads every part of a deck
 * the user picked. `docs/DEPENDENCY-ALERTS.md` has the reading. Nothing was
 * ever exposed, because `npm ci` installs the lockfile — which is exactly why
 * nothing caught it either.
 *
 * **What this checks, precisely, because the file's name is wider than its
 * reach**: that every declared floor IS the version the lockfile resolved. It
 * does NOT ask whether that version has advisories against it. It cannot: the
 * advisory database changes daily and reaching it needs the network, so a test
 * built on it would go red on a day nobody touched the code — and this project
 * keeps that kind of check out of `test` on purpose, the way `ci.yml` keeps
 * Microsoft's manifest validator out of it.
 *
 * What the alignment buys instead is a rule that needs no network and cannot
 * rot: **the floor is the version that was tested.** Whatever `npm audit`,
 * `npm ci` and the whole suite have said about the installed tree, they have
 * now said about the lowest version this manifest permits, because those are
 * the same version. A floor below the lock is a version nothing here has ever
 * run.
 *
 * The rules live in this file rather than inline in the test so there is one
 * copy for the suite and for any tool that wants the same answer — the split
 * `scripts/package-integrity.mjs` and `scripts/coverage-scope.mjs` already use.
 */

/** Where npm records an installed package inside `package-lock.json`. */
const LOCK_PREFIX = "node_modules/";

/**
 * The lowest version a range admits, or `null` if that cannot be read.
 *
 * Only the three shapes npm writes by itself are understood: `^1.2.3`, `~1.2.3`
 * and a bare `1.2.3`. Everything else — `*`, `>=1`, a union, a URL, `file:`,
 * `workspace:` — answers `null`, and the caller REFUSES it rather than passing
 * it. A scanner that quietly skips what it cannot parse is the shape of guard
 * this repository has caught twice: it answers "clean" about a question it
 * never asked. If a range like that is ever wanted here, this function is where
 * the decision gets made and written down.
 */
export function floorOf(range) {
  if (typeof range !== "string") return null;
  const match = /^[~^]?(\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?)$/.exec(range.trim());
  return match ? match[1] : null;
}

/**
 * Every direct dependency's declared range, in one list.
 *
 * `dependencies` and `devDependencies` both, because the hazard does not care
 * which: `vitest` is a development dependency and its floor named a version its
 * own companion refused to install beside.
 */
function declared(manifest) {
  return [
    ...Object.entries(manifest.dependencies ?? {}).map(([name, range]) => ({ name, range, dev: false })),
    ...Object.entries(manifest.devDependencies ?? {}).map(([name, range]) => ({ name, range, dev: true })),
  ];
}

/**
 * What is wrong with the floors, as a list of sentences a reader can act on.
 *
 * An empty list means every declared floor is the installed version. The count
 * of what was SCANNED is returned beside it, because an empty problem list is
 * also what a scanner that has stopped matching returns — the vacuity the
 * docProps guard was open to for weeks. A caller that does not check
 * `scanned` has not checked anything.
 */
export function floorProblems({ manifest, lock }) {
  const problems = [];
  const entries = declared(manifest);
  const root = lock?.packages?.[""] ?? {};

  for (const { name, range, dev } of entries) {
    const where = dev ? "devDependencies" : "dependencies";

    const floor = floorOf(range);
    if (floor === null) {
      problems.push(
        `${name} is declared as "${range}" in ${where}, and this check cannot read a floor from that. ` +
          `Only ^x.y.z, ~x.y.z and x.y.z are understood. Refusing rather than passing it.`,
      );
      continue;
    }

    const locked = lock?.packages?.[LOCK_PREFIX + name]?.version;
    if (typeof locked !== "string") {
      problems.push(`${name} is declared in ${where} but the lockfile records no installed version for it.`);
      continue;
    }

    if (floor !== locked) {
      problems.push(
        `${name} declares a floor of ${floor} in ${where} but the lockfile installs ${locked}. ` +
          `Nothing here has ever run ${floor}: raise the range to "^${locked}" (or re-run npm install).`,
      );
    }

    // The lockfile keeps its own copy of the root ranges, and `npm install` is
    // what makes the two agree. Editing `package.json` by hand and committing
    // without it leaves a manifest whose ranges the lockfile contradicts — and
    // the check above would still pass, because it never reads this copy.
    const mirrored = root[where]?.[name];
    if (mirrored !== undefined && mirrored !== range) {
      problems.push(
        `${name} is "${range}" in package.json but "${mirrored}" in the lockfile's own ${where}. ` +
          `Run npm install so the two agree.`,
      );
    }
  }

  return { problems, scanned: entries.length };
}
