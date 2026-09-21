#!/usr/bin/env node
/**
 * Audit the versions this project's ranges would ACCEPT, not the ones it installed.
 *
 * `npm audit` reads the lockfile, so it answers about one version per package:
 * the one that resolved. A caret is a promise about a whole interval, and its
 * lower bound is the version a fresh `npm install` without a lockfile — or any
 * consumer resolving these ranges — is entitled to pick. Nothing looked at that
 * bound until 2026-09-21, when `@xmldom/xmldom` turned out to be declared
 * `^0.9.8` with 17 advisories against 0.9.8, one of them high severity, behind
 * the parser that reads every part of a user-supplied deck. Every check on the
 * repository was green, because the lockfile held 0.9.12 and that is what
 * `npm audit` had been reading. `docs/DEPENDENCY-ALERTS.md` has the reading.
 *
 * `test/floors.test.ts` holds the half of this that needs no network: that each
 * declared floor IS the installed version, so the floor is a version the whole
 * gate has already run. This script is the other half and it cannot be a test —
 * it needs the registry and the advisory database, and the advisory database
 * changes daily, so a required check built on it would go red on a day nobody
 * touched the code. It runs weekly instead, in the shape `pane-audit.yml` and
 * `sibling-watch.yml` already use.
 *
 * Usage:
 *   node scripts/audit-floors.mjs            # audit this repo's floors
 *   node scripts/audit-floors.mjs --json     # machine-readable output
 *
 * Exit 0 when every floor is clean, 3 when a floor carries advisories, and
 * anything else when the SWEEP broke. A broken sweep must never read as a clean
 * one, which is the failure mode this whole file is shaped around.
 *
 * **It runs a control arm first, and refuses to report at all if the control
 * comes back clean.** `scripts/mutate-core.mjs` is the cautionary tale: it told
 * its reader to check that the unmutated copy was green and did not check, the
 * copy was red for an unrelated reason, every mutant read "caught", and the run
 * ended on a line that looked like a perfect score. A sweep whose failure mode
 * is a GOOD-looking result has to measure its own baseline. Here the baseline
 * is a package-version pair known to carry advisories — `@xmldom/xmldom@0.9.8`,
 * 17 of them as measured on 2026-09-21 — audited in its own scratch tree. If
 * the registry is unreachable, the advisory endpoint is down, or the JSON
 * schema moves, the control reports nothing and this script exits non-zero
 * WITHOUT looking at the real floors.
 *
 * The control's own rot is in the safe direction, which is why it is allowed to
 * be a fixed version: if those advisories are ever withdrawn the control goes
 * quiet, the sweep refuses, and somebody reads the failure. It cannot rot into
 * silence.
 */
import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { floorOf } from "./dependency-floors.mjs";
import { isMain } from "./is-main.mjs";

/**
 * The pair the control arm asks about, and when its answer was measured.
 *
 * Undated, this is a claim that rots invisibly. Dated, it is a recording: the
 * number can only get older, and a reader can judge that for themselves.
 */
export const CONTROL = { name: "@xmldom/xmldom", version: "0.9.8", advisoriesOn: "2026-09-21", wereAtLeast: 10 };

/**
 * The scratch manifest: every declared range pinned to its own lower bound.
 *
 * Pinned EXACTLY, with no caret, because the question is about that version and
 * nothing above it. A range whose floor cannot be read is returned in
 * `unreadable` rather than dropped — `test/floors.test.ts` already fails the
 * build for one, so reaching this with any is itself a finding, and a sweep
 * that silently audited fewer packages than the manifest declares would be
 * answering a narrower question than its name.
 */
export function floorManifest(manifest) {
  const pinned = {};
  const unreadable = [];
  const declared = { ...(manifest.dependencies ?? {}), ...(manifest.devDependencies ?? {}) };

  for (const [name, range] of Object.entries(declared)) {
    const floor = floorOf(range);
    if (floor === null) unreadable.push({ name, range });
    else pinned[name] = floor;
  }

  return { pinned, unreadable };
}

/**
 * What `npm audit --json` said, or why it cannot be believed.
 *
 * Every refusal here is a shape this script would otherwise report as "no
 * advisories". A report with no `metadata`, a report from a schema this does
 * not know, a count that is not a number: each is indistinguishable from a
 * clean answer if it is read carelessly, and each is the reason the caller
 * treats `problems` as fatal rather than cosmetic.
 */
export function advisoryReport(audit) {
  const problems = [];
  if (!audit || typeof audit !== "object") {
    return { findings: [], counted: null, problems: ["npm audit produced no readable JSON at all."] };
  }
  if (audit.auditReportVersion !== 2) {
    problems.push(
      `npm audit reported schema version ${JSON.stringify(audit.auditReportVersion)}, and this sweep reads version 2. ` +
        `Refusing to read it rather than reporting a clean sweep from a shape it does not understand.`,
    );
  }

  const counted = audit.metadata?.vulnerabilities;
  if (!counted || typeof counted.total !== "number") {
    problems.push("npm audit reported no vulnerability totals, so there is nothing to compare against.");
  }

  const findings = [];
  for (const [name, entry] of Object.entries(audit.vulnerabilities ?? {})) {
    const advisories = (entry?.via ?? []).filter((via) => typeof via === "object" && via !== null);
    findings.push({
      name,
      severity: entry?.severity ?? "unknown",
      direct: entry?.isDirect === true,
      advisories: advisories.length,
      titles: advisories.map((via) => String(via.title ?? via.url ?? "(untitled advisory)")),
    });
  }

  return { findings, counted: counted ?? null, problems };
}

/**
 * Resolve and audit one set of exact pins, in a directory of its own.
 *
 * `--package-lock-only` because nothing here needs the packages on disk: the
 * advisory question is answered from the resolved tree, and installing six
 * hundred files to ask it would make a weekly job slow enough to switch off.
 *
 * The resolved lockfile is read back and checked against the pins before the
 * audit is believed. npm is free to resolve something other than what was
 * asked for — a pin npm cannot satisfy, a registry that redirects — and an
 * audit of a tree that is not the one under test is the quietest possible
 * wrong answer.
 */
export function auditPins(pins, { run = npmIn } = {}) {
  const dir = mkdtempSync(join(tmpdir(), "ssf-floor-audit-"));
  writeFileSync(
    join(dir, "package.json"),
    `${JSON.stringify({ name: "floor-audit", version: "1.0.0", private: true, dependencies: pins }, null, 2)}\n`,
  );

  const problems = [];
  try {
    run(dir, ["install", "--package-lock-only", "--no-audit", "--no-fund", "--silent"]);
  } catch (err) {
    return { findings: [], counted: null, problems: [`the scratch resolve failed: ${messageOf(err)}`], dir };
  }

  let lock;
  try {
    lock = JSON.parse(readFileSync(join(dir, "package-lock.json"), "utf8"));
  } catch (err) {
    return {
      findings: [],
      counted: null,
      problems: [`the scratch lockfile could not be read: ${messageOf(err)}`],
      dir,
    };
  }

  for (const [name, version] of Object.entries(pins)) {
    const resolved = lock.packages?.[`node_modules/${name}`]?.version;
    if (resolved !== version) {
      problems.push(
        `asked for ${name}@${version} and the scratch tree resolved ${resolved ?? "nothing"}. ` +
          `Refusing: an audit of a tree that is not the one under test answers about the wrong versions.`,
      );
    }
  }

  // `npm audit` exits non-zero WHEN IT FINDS SOMETHING, which is the answer
  // rather than an error, so the status is ignored and the JSON is what is
  // read. A genuinely broken run fails to produce readable JSON, and
  // `advisoryReport` refuses that.
  let raw;
  try {
    raw = run(dir, ["audit", "--json"], { tolerateFailure: true });
  } catch (err) {
    return { findings: [], counted: null, problems: [`npm audit could not be run: ${messageOf(err)}`], dir };
  }

  let parsed = null;
  try {
    parsed = JSON.parse(raw);
  } catch {
    /* advisoryReport turns this into a refusal, with the same wording as every other unreadable shape */
  }

  const report = advisoryReport(parsed);
  return { ...report, problems: [...problems, ...report.problems], dir };
}

function messageOf(err) {
  return err instanceof Error ? err.message : String(err);
}

function npmIn(dir, args, { tolerateFailure = false } = {}) {
  try {
    return execFileSync("npm", args, { cwd: dir, encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] });
  } catch (err) {
    // npm puts the report on stdout even when it exits 1 for having found
    // something. Only a run with nothing on stdout is actually a failure.
    const stdout = err !== null && typeof err === "object" && "stdout" in err ? err.stdout : null;
    if (tolerateFailure && typeof stdout === "string" && stdout.length > 0) return stdout;
    throw err;
  }
}

/** The report a human reads, in the shape `sibling-watch.mjs` writes one. */
export function report({ findings, unreadable, counted, control }) {
  const lines = [];
  lines.push("# Floor audit");
  lines.push("");
  lines.push(
    `The versions this project's ranges would ACCEPT, audited at their lower bounds. ` +
      `\`npm audit\` on the lockfile answers about what resolved; this answers about what a resolve is allowed to pick.`,
  );
  lines.push("");
  lines.push(
    `Control: \`${CONTROL.name}@${CONTROL.version}\` reported ${control} advisories ` +
      `(at least ${CONTROL.wereAtLeast} expected, measured ${CONTROL.advisoriesOn}). The sweep can see advisories.`,
  );
  lines.push("");

  if (unreadable.length) {
    lines.push("## Ranges with no readable floor");
    lines.push("");
    for (const { name, range } of unreadable)
      lines.push(`- \`${name}\` is declared \`${range}\`, and was NOT audited.`);
    lines.push("");
  }

  if (!findings.length) {
    lines.push(`No advisories against any declared floor. npm audit counted ${counted?.total ?? 0} in total.`);
    return `${lines.join("\n")}\n`;
  }

  lines.push("## Floors carrying advisories");
  lines.push("");
  for (const finding of findings) {
    lines.push(
      `### \`${finding.name}\` — ${finding.severity}, ${finding.advisories} ${finding.advisories === 1 ? "advisory" : "advisories"}${finding.direct ? ", a direct dependency" : ""}`,
    );
    lines.push("");
    for (const title of finding.titles) lines.push(`- ${title}`);
    lines.push("");
  }
  lines.push(
    "The remedy is to raise the range so its floor is a version without these, then record the reading in " +
      "`docs/DEPENDENCY-ALERTS.md`. Nothing is necessarily exposed — `npm ci` installs the lockfile — but the manifest " +
      "is promising something it should not.",
  );
  return `${lines.join("\n")}\n`;
}

export function main(
  argv = [],
  { read = readFileSync, audit = auditPins, out = process.stdout, err = process.stderr } = {},
) {
  const json = argv.includes("--json");

  // The control arm, first and on its own. If this cannot see 17 advisories
  // against a version that has them, nothing this script says about the real
  // floors is worth reading.
  const control = audit({ [CONTROL.name]: CONTROL.version });
  const controlAdvisories = control.findings.find((f) => f.name === CONTROL.name)?.advisories ?? 0;
  if (control.problems.length || controlAdvisories < CONTROL.wereAtLeast) {
    err.write(
      `audit-floors: the control arm failed, so the sweep will not report.\n` +
        `  ${CONTROL.name}@${CONTROL.version} reported ${controlAdvisories} advisories; at least ${CONTROL.wereAtLeast} were there on ${CONTROL.advisoriesOn}.\n` +
        control.problems.map((p) => `  ${p}\n`).join(""),
    );
    return 1;
  }

  const manifest = JSON.parse(read("package.json", "utf8"));
  const { pinned, unreadable } = floorManifest(manifest);
  const result = audit(pinned);

  if (result.problems.length) {
    err.write(`audit-floors: the sweep broke.\n${result.problems.map((p) => `  ${p}\n`).join("")}`);
    return 1;
  }

  out.write(
    json
      ? `${JSON.stringify({ control: controlAdvisories, pinned, unreadable, findings: result.findings }, null, 2)}\n`
      : report({ ...result, unreadable, control: controlAdvisories }),
  );

  return result.findings.length || unreadable.length ? 3 : 0;
}

if (isMain(import.meta.url)) {
  try {
    process.exit(main(process.argv.slice(2)));
  } catch (err) {
    process.stderr.write(`audit-floors: ${messageOf(err)}\n`);
    process.exit(1);
  }
}
