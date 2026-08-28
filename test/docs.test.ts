import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { STEP_TITLE } from "../src/pane/steps.js";

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

/**
 * Every picture format the engine answers to, taken from its own table.
 *
 * A separate reader from `formatKinds` because the image formats are decided in
 * `merge/images.ts` rather than in the format switch — a manual that documented
 * only the switch would be complete and still miss half the formats a user can
 * write.
 */
function imageFormats(): string[] {
  const src = readFileSync("src/core/merge/images.ts", "utf8");
  const table = src.match(/const MODES[^=]*=\s*\{([^}]+)\}/)?.[1] ?? "";
  return [...table.matchAll(/"?([a-z-]+)"?\s*:/g)].map((m) => m[1] ?? "");
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

  it("documents every picture format the engine accepts", () => {
    const kinds = imageFormats();
    // Same vacuity guard as the format switch above: a regex that stops
    // matching would otherwise pass this test forever.
    expect(kinds).toContain("image");
    expect(kinds.length).toBeGreaterThan(2);
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

  it("quotes button labels that still exist in the pane", () => {
    /**
     * The manual's quickstart walks somebody through their first merge by
     * naming the buttons. Three of those labels were wrong in the first draft,
     * written from memory: the preview button is "Preview the first row" and
     * not "Press it", and putting the row back is "Remove the preview" and not
     * "Put it back". A walkthrough naming a button that is not there is worse
     * than no walkthrough — the reader concludes the add-in is broken.
     *
     * Only the STATIC labels. The three that carry a number
     * (`Use slides 3 to 3`, `Use 3 rows`, `Add 3 slides`) are built from the
     * state, and pinning their template here would break on a whitespace
     * change while catching nothing a rename would not already trip below.
     */
    const source = readFileSync("src/pane/steps.ts", "utf8") + readFileSync("src/pane/render.ts", "utf8");
    // Whitespace-collapsed, because the manual is wrapped: the formatter breaks
    // "Choose the slides that repeat" across two lines and an exact match would
    // fail on a correct document — the false red this repo has a module about.
    const prose = manual.replace(/\s+/g, " ");
    for (const label of [
      "Choose the slides that repeat",
      "Preview the first row",
      "Remove the preview",
      "Skip to the merge",
      // The two controls the five-step order added. Both are the answer to
      // "how do I insert the fields?", which is the question a first run
      // actually asked, so a manual that stops naming them is a manual that
      // has lost the answer.
      "Check the slides for fields",
      "Click into a text box on the slide",
      "Remove these slides",
      "What this run did, step by step",
      // The condition control's shut line, which is the only thing naming the
      // feature on screen.
      "Every slide, every row",
    ]) {
      expect(source, `the pane no longer has a "${label}" button`).toContain(label);
      expect(prose, `the manual does not mention "${label}"`).toContain(label);
    }
  });

  it("names every step the pane actually renders", () => {
    for (const title of Object.values(STEP_TITLE)) {
      expect(manual, `step ${title} is not in the manual`).toContain(title);
    }
  });

  it("does not call the pane planned, because the pane is built", () => {
    /**
     * The direction the test above cannot see. It asserts the word "planned"
     * APPEARS, which stays true forever and is satisfied by a manual describing
     * shipped work as unbuilt — which is what happened: the status block said
     * "the task pane is not written yet" for days after the pane shipped, on
     * the first screen of the document somebody reads before installing.
     *
     * Claiming less than you have is not the harmless direction. It tells a
     * reader not to look for the thing that is there.
     */
    const section = manual.slice(manual.indexOf("## The pane"));
    const body = section.slice(0, section.indexOf("\n## ", 3));
    expect(body.toLowerCase(), "the pane section is marked planned").not.toContain("planned");
    expect(manual).not.toContain("the task pane is not written");
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
