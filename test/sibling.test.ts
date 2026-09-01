import { readdirSync, readFileSync, statSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";

// @ts-expect-error — plain .mjs with no types, and deliberately the SAME module
// the weekly sweep runs. Two copies of the triage table is how a claim quietly
// stops matching its check.
import { TRIAGED, VERDICTS, tableKeys, tablesFrom, untriaged, verdictOf } from "../scripts/sibling-watch.mjs";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import { withoutTsProse } from "../scripts/without-prose.mjs";

/**
 * The borrowed-fact guard.
 *
 * Most of what this repo knows about the PowerPoint host was learned by a
 * sibling project and hand-copied here — 44 citations across 24 files when they
 * were first counted. `docs/SIBLING.md` is the ledger; this is the half of it a
 * machine can hold.
 *
 * What actually rotted was narrow. A single run's observation ("a by-id
 * clean-up reported 45 deletes and removed nothing") is true forever and needs
 * nothing. A COUNT OF ROUNDS is a live counter by definition: the sibling runs
 * more of them. Four source comments said "174 consecutive archived rounds" and
 * one said "passed 174 of 174" — all correct on the morning they were written,
 * all wrong the moment round 175 ran, with nothing anywhere to say so.
 *
 * So: a round count carries the date it was taken, or it is not a round count.
 * Dated, the number stops being a claim and becomes a recording — it never
 * becomes false, only older, and a reader can judge that for themselves.
 *
 * **Scoped to source and scripts, never to docs, and that is deliberate.**
 * `docs/SIBLING.md` and `CLAUDE.md` state the rule, which means they quote the
 * exact sentences that break it. A guard that read those as claims would fail
 * on the two files explaining why it exists — the "a guard that goes red for
 * the wrong reason is worse than no guard" failure, which this repo has already
 * recorded once when the no-Office-imports test matched the word "Office.js" in
 * the comments explaining why the engine avoids it. Source comments are what
 * justify code, and they are what went stale.
 */

/**
 * This file, which is the one place in `test/` whose SUBJECT is the rule.
 *
 * Excluded by name, not by an opt-out marker anyone could reach for. A guard
 * that can be silenced from the file it is guarding is not a guard. This file
 * quotes the sentences that broke the rule so a reader can see the shape, and
 * on its first run it duly failed on its own prose — the same failure as the
 * no-Office-imports test matching "Office.js" in the comments explaining why
 * the engine avoids it. Docs are excluded for the same reason and only that
 * reason.
 *
 * Built with `join` rather than written as `"test/sibling.test.ts"`, because it
 * is compared against a path `join` produced. On Windows that literal never
 * matches `test\sibling.test.ts`, so this file scanned itself and failed on the
 * two sentences it quotes to show the shape — a guard going red for the wrong
 * reason, on the owner's own box, while CI on ubuntu stayed green.
 */
const SELF = join("test", "sibling.test.ts");

/** Every .ts and .mjs file under the directories that hold justifying comments. */
function sourceFiles(): string[] {
  const roots = ["src", "scripts", "test", "probe"];
  const out: string[] = [];
  const walk = (dir: string): void => {
    for (const name of readdirSync(dir)) {
      const path = join(dir, name);
      if (statSync(path).isDirectory()) walk(path);
      else if (/\.(ts|mjs)$/.test(path) && path !== SELF) out.push(path);
    }
  };
  for (const root of roots) walk(root);
  return out;
}

/**
 * A file as one string, with comment prefixes stripped, plus where each line
 * began.
 *
 * A line-by-line scan CANNOT SEE THE DEFECT THIS TEST WAS WRITTEN FOR. The
 * sentence in `powerpoint.ts` wrapped as `…across **174 consecutive` /
 * `* archived rounds every rung answered**`, so the claim existed only across
 * the join and every line on its own was innocent. The first version of this
 * test scanned lines, passed against the unfixed file, and was decoration.
 *
 * Offsets are kept so a match can still be reported at the line it starts on —
 * naming the line is the whole value of the failure message.
 */
function joined(text: string): { body: string; starts: number[] } {
  const lines = text.split("\n").map((l) => l.replace(/^\s*(?:\*|\/\/)\s?/, ""));
  const starts: number[] = [];
  let at = 0;
  for (const line of lines) {
    starts.push(at);
    at += line.length + 1;
  }
  return { body: lines.join(" "), starts };
}

/** Which 1-based line an offset into `joined`'s body falls on. */
function lineAt(starts: number[], offset: number): number {
  let lo = 0;
  for (let i = 0; i < starts.length; i++) if ((starts[i] ?? 0) <= offset) lo = i;
  return lo + 1;
}

/**
 * A claim about how many ROUNDS the sibling has run.
 *
 * Deliberately not "any number near the word sibling". `45 successful deletes`,
 * `46 InvalidParam`, `37 generated slides` and `a 680-second run` are all
 * single observations that will read correctly in ten years; flagging them
 * would train everyone to ignore this test. A round count is the one shape that
 * is guaranteed to move, because the sibling's whole method is running more.
 *
 * `one round` and `nine rounds` spelled as words are caught too — the digits
 * are not what makes it a counter.
 */
const ROUND_COUNT =
  /\b(\d+|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|seventeen|twenty)\s+(?:consecutive\s+)?(?:archived\s+)?(?:real-host\s+)?rounds\b/gi;

/** An ISO date, which is how this repo stamps a measurement. */
const DATED = /\b20\d\d-\d\d-\d\d\b/;

/** Lines either side of a claim that may carry its date. */
const WINDOW = 6;

describe("a borrowed round count carries its date", () => {
  it("finds no undated round count in any source comment", () => {
    const offenders: string[] = [];
    for (const file of sourceFiles()) {
      const text = readFileSync(file, "utf8");
      const lines = text.split("\n");
      const { body, starts } = joined(text);
      ROUND_COUNT.lastIndex = 0;
      for (let m = ROUND_COUNT.exec(body); m; m = ROUND_COUNT.exec(body)) {
        const line = lineAt(starts, m.index);
        const near = lines.slice(Math.max(0, line - 1 - WINDOW), line + WINDOW).join("\n");
        if (DATED.test(near)) continue;
        offenders.push(`${file}:${line}  …${m[0]}…`);
      }
    }
    // Named rather than counted: "3 undated claims" sends the reader hunting,
    // and the whole point of the rule is that a stale number is invisible.
    expect(offenders, `undated round counts (see docs/SIBLING.md):\n${offenders.join("\n")}`).toEqual([]);
  });

  it("catches the shape that actually rotted", () => {
    // The comment that was in `powerpoint.ts` until 2026-08-27, verbatim,
    // WRAPPED AS IT WAS. Run through `joined` because that is what the scan
    // does, and because a line-by-line version of this test passed against the
    // unfixed file — the claim lives only across the join.
    const rotted = [
      " * in every round of its self-test battery, and across **174 consecutive",
      " * archived rounds every rung answered**, in 550-710ms, with zero refusals",
    ].join("\n");
    ROUND_COUNT.lastIndex = 0;
    expect(ROUND_COUNT.test(joined(rotted).body)).toBe(true);
    expect(DATED.test(rotted)).toBe(false);

    // And the fix passes: the same claim with the date it was measured on.
    const fixed = rotted + "\n * and zero silences. Measured 2026-08-27.";
    expect(DATED.test(fixed)).toBe(true);
  });

  it("leaves a single run's observation alone", () => {
    // Every one of these is durable: it happened once and will read correctly
    // forever. Flagging them is how this test would become noise.
    for (const durable of [
      "a sibling project's by-id clean-up once reported 45 successful deletes",
      "a sibling add-in logged 46 `InvalidParam passed to GetItem(id)` failures in one run",
      "a sibling project put 37 generated slides ahead of somebody's title slide",
      "logged a 680-second run that shipped duplicate slides",
      "a sibling add-in pages every collection read at 20 for that reason",
    ]) {
      ROUND_COUNT.lastIndex = 0;
      expect(ROUND_COUNT.test(durable), durable).toBe(false);
    }
  });
});

describe("the ledger is wired in", () => {
  const ledger = readFileSync("docs/SIBLING.md", "utf8");

  it("is reachable from the places somebody starts reading", () => {
    for (const [file, why] of [
      ["CLAUDE.md", "the memory file's host rules"],
      ["README.md", "the docs table"],
      ["CONTRIBUTING.md", "the note before touching PowerPoint"],
    ] as const) {
      expect(readFileSync(file, "utf8"), why).toContain("docs/SIBLING.md");
    }
  });

  it("answers every row, including the ones that are no exposure", () => {
    // "No exposure" is the answer nobody writes down, and an untriaged finding
    // is indistinguishable from an unnoticed one. If the ledger ever stops
    // carrying them it has become a list of work instead of a record of
    // decisions.
    expect(ledger).toMatch(/no exposure/i);
  });

  it("says out loud that reading it by hand needs both repositories", () => {
    // The precondition that made this transferable at all, and the one a
    // session holding only this repo cannot discover for itself.
    expect(ledger).toMatch(/both repositories checked out in the same session/i);
  });
});

describe("the sibling sweep", () => {
  const table = readFileSync("scripts/sibling-watch.mjs", "utf8");

  it("pulls the keys out of a table, quoted or bare", () => {
    const src = [
      "export const T = {",
      '  bare: "a",',
      '  "quoted": "b",',
      '  "with-dashes": "c",',
      // An id spelled with an underscore or a dot was DROPPED, silently, from a
      // sweep whose whole job is noticing something new. The sibling's ids
      // happen to use hyphens today, which is what made it latent rather than
      // broken.
      '  under_scored: "d",',
      '  "dotted.id": "e",',
      "};",
      "",
    ].join("\n");
    expect(tableKeys(src, "T")).toEqual(["bare", "quoted", "with-dashes", "under_scored", "dotted.id"]);
  });

  it("says a missing table is MISSING, never empty", () => {
    // The failure this sweep would otherwise commit against itself: a table
    // renamed upstream would match nothing, report "nothing new" every Monday,
    // and be indistinguishable from a quiet week forever.
    expect(tableKeys("export const OTHER = {\n  a: 1,\n};\n", "GONE")).toBeNull();
    expect(() => {
      tablesFrom((): string => "export const NOTHING = {\n};\n");
    }).toThrow(/renamed, moved, emptied or reformatted/);
  });

  it("says a table it can no longer READ is missing too", () => {
    /**
     * The same silence by a likelier route. The key pattern is anchored on
     * two-space indentation, so a reindent upstream — a formatter run, one more
     * level of nesting — parses to an empty array, which is not null and was
     * therefore reported as a quiet week. The protection above covered a rename
     * and not a reformat.
     *
     * A curated table that genuinely holds nothing is not worth distinguishing
     * here: saying so once costs a person a minute, and reading it as quiet
     * costs every Monday after.
     */
    const reindented = ["export const T = {", '    bare: "a",', '    "with-dashes": "c",', "};", ""].join("\n");
    expect(tableKeys(reindented, "T"), "four spaces is a table this cannot read").toBeNull();
    expect(() => {
      tablesFrom((): string => reindented.replace("T", "FAKE_BASELINE"));
    }).toThrow(/emptied or reformatted/);
  });

  it("reports a finding with no row, and keeps which tables it was in", () => {
    const tables = [
      { kind: "question", table: "FAKE_BASELINE", keys: ["known", "fresh"] },
      { kind: "question", table: "UNSTABLE_ANSWERS", keys: ["fresh"] },
    ];
    const found = untriaged(tables, { "question:known": "NO EXPOSURE — nothing here." });
    // One row, not two: the same finding in two tables is one thing to answer.
    // Which tables it sits in rides along, because UNSTABLE_ANSWERS means the
    // sibling has seen it answer two ways and nothing may be built on it.
    expect(found).toEqual([
      { id: "question:fresh", kind: "question", key: "fresh", tables: ["FAKE_BASELINE", "UNSTABLE_ANSWERS"] },
    ]);
  });

  it("cannot be fooled by a key inherited from Object.prototype", () => {
    // `TRIAGED` is looked up by a key that comes from a file fetched over the
    // network, so a bare `triaged[id]` would answer truthy for `constructor`
    // and swallow a real finding in silence.
    //
    // TWO things make this true and the test cannot tell them apart: the
    // `kind:` prefix (so the id is never a bare prototype name) and the
    // `hasOwnProperty` lookup. Removing EITHER alone leaves this passing —
    // measured, by doing it — and removing both turns it red. Stated rather
    // than dressed up as proof of one guard, because a test that passes
    // against the unfixed file is decoration and this one nearly was.
    const found = untriaged([{ kind: "question", table: "T", keys: ["constructor", "__proto__"] }], {});
    expect(found.map((f: { key: string }) => f.key)).toEqual(["constructor", "__proto__"]);
  });

  it("opens every reason with a verdict from the closed vocabulary", () => {
    const unparsed = Object.entries(TRIAGED).filter(([, why]) => verdictOf(why) === undefined);
    expect(
      unparsed.map(([k]) => k),
      "a reason must open with one of " + VERDICTS.join(", "),
    ).toEqual([]);
  });

  it("gives every finding a reason, not just a verdict", () => {
    // "NO EXPOSURE" alone is a shrug, and a key added purely to silence the
    // sweep is worse than the finding it silences.
    //
    // The bar is deliberately low. A first pass required twenty characters and
    // failed on four grouping entries reading "no grouping." — which is a
    // COMPLETE answer for a finding about grouping. Padding those to satisfy a
    // threshold would have made the table worse to read in the name of a test,
    // so the threshold moved instead.
    const bare = Object.entries(TRIAGED as Record<string, string>).filter(
      ([, why]) => why.replace(verdictOf(why) ?? "", "").replace(/^[\s—-]+/, "").length < 8,
    );
    expect(bare.map(([k]) => k)).toEqual([]);
  });

  it("keys everything as issue: or question:", () => {
    // The sibling's own spelling, so a rename upstream surfaces as a new
    // finding rather than quietly matching nothing.
    const odd = Object.keys(TRIAGED).filter((k) => !/^(issue:\d+|question:[a-z0-9-]+)$/.test(k));
    expect(odd).toEqual([]);
  });

  it("keeps the ledger from falling behind the table", () => {
    // The drift this whole apparatus exists to stop, applied to itself. Every
    // finding we ACTED on has to appear in the prose a human reads; the ones
    // that are no exposure do not, or the ledger becomes a list of 52 shrugs.
    // This found 13 missing rows the first time it ran, against a ledger
    // written the same morning.
    const ledger = readFileSync("docs/SIBLING.md", "utf8");
    const missing = Object.entries(TRIAGED as Record<string, string>)
      .filter(([, why]) => verdictOf(why) !== "NO EXPOSURE")
      .map(([key]) => key)
      .filter((key) => !ledger.includes(key.split(":")[1] ?? ""));
    expect(missing, "acted on but absent from docs/SIBLING.md").toEqual([]);
  });

  it("reads raw files and never runs the sibling's code", () => {
    // A weekly job that imports a file fetched over the network is a supply
    // chain, not a sweep.
    expect(table).toContain("raw.githubusercontent.com");
    expect(withoutTsProse(table)).not.toMatch(/\bimport\s*\(/);
  });
});
