import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

/**
 * The lockstep guard.
 *
 * Documentation that is updated when someone remembers is documentation that is
 * wrong, and stale docs are worse than none: a reader trusts them. These tests
 * read the surfaces out of the SOURCE and fail when the docs have not kept up,
 * so a feature and its documentation land in the same change or neither does.
 */

const manual = readFileSync("docs/MANUAL.md", "utf8");
const readme = readFileSync("README.md", "utf8");
const changelog = readFileSync("CHANGELOG.md", "utf8");
const backlog = readFileSync("docs/BACKLOG.md", "utf8");

/** Every format kind `applyFormat` answers to, taken from its switch. */
function formatKinds(): string[] {
  const src = readFileSync("src/core/data/format.ts", "utf8");
  return [...src.matchAll(/case "([a-z]+)":/g)].map((m) => m[1] ?? "");
}

/** Every tag key the engine writes, taken from its exported constants. */
function tagKeys(): string[] {
  const src = readFileSync("src/core/pptx/tags.ts", "utf8");
  return [...src.matchAll(/export const TAG_\w+ = "([^"]+)"/g)].map((m) => m[1] ?? "");
}

describe("the manual keeps up with the code", () => {
  it("documents every format the engine accepts", () => {
    const kinds = formatKinds();
    // A guard that measures nothing passes forever. If the switch stops
    // matching, this fails here rather than pretending every format is
    // documented.
    expect(kinds.length).toBeGreaterThan(3);
    for (const kind of kinds) expect(manual, `format "${kind}" is not in the manual`).toContain(`|${kind}`);
  });

  it("documents every tag the engine writes", () => {
    const keys = tagKeys();
    expect(keys.length).toBeGreaterThan(3);
    for (const key of keys) expect(manual, `tag ${key} is not in the manual`).toContain(key);
  });

  it("says which parts are not built yet, rather than describing them as shipped", () => {
    // The manual documents a design that is ahead of the code. That is fine as
    // long as it never claims to be behind it.
    expect(manual).toContain("planned");
  });
});

describe("the documentation set is whole", () => {
  it("keeps an Unreleased section in the changelog", () => {
    expect(changelog).toContain("## [Unreleased]");
  });

  it("keeps a rejected list in the backlog, so the same idea is not re-proposed", () => {
    expect(backlog).toContain("Rejected");
  });

  it("points at every document from the README", () => {
    for (const doc of ["docs/MANUAL.md", "docs/BACKLOG.md", "CHANGELOG.md"]) {
      expect(readme, `${doc} is not linked from the README`).toContain(doc);
    }
  });
});

describe("the probe's question numbers", () => {
  /**
   * The table in `docs/PROBE.md` and the headings `read-answers.mjs` prints are
   * two hand-maintained lists of the same questions, and they had already come
   * apart: the export arm was added as question 5 in the reader while the doc
   * still called `fill.setImage` 5, so the same number named two different
   * questions and a sheet read against the doc answered the wrong one.
   *
   * The NUMBERS are pinned and the wording deliberately is not. A gate over the
   * text would make every rephrasing a two-file edit and would be switched off
   * within a month; the drift that actually happened was numeric.
   */
  const doc = readFileSync("docs/PROBE.md", "utf8");
  const reader = readFileSync("scripts/read-answers.mjs", "utf8");
  const inDoc = [...doc.matchAll(/^\| (\d+) \| /gm)].map((m) => Number(m[1]));
  // The reader prints them as "\n5. Does …" at the head of each section.
  const inReader = [...reader.matchAll(/"\\n(\d+)\. /g)].map((m) => Number(m[1]));

  it("finds both lists", () => {
    // Guards the two patterns. Either one matching nothing would make the
    // comparison below pass on a pair of empty arrays.
    expect(inDoc.length).toBeGreaterThan(5);
    expect(inReader.length).toBeGreaterThan(5);
  });

  it("asks each number once, in order, in both places", () => {
    expect(inDoc).toStrictEqual([...inDoc].sort((a, b) => a - b));
    expect(inReader).toStrictEqual([...inReader].sort((a, b) => a - b));
    expect(new Set(inDoc).size).toBe(inDoc.length);
    // Question 0 is the control arm, which the reader folds into section 1
    // rather than printing on its own line.
    expect(inReader).toStrictEqual(inDoc.filter((n) => n !== 0));
  });
});
