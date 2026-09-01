/**
 * No deck this repository carries names the person who authored it.
 *
 * Every .pptx here was made by a human in PowerPoint or by python-pptx, on
 * purpose — the test kit's whole argument is that the engine must be tested
 * against parts it did not write itself. PowerPoint stamps whoever saved the
 * file into `docProps`, so the price of that argument is a real name in a
 * public repository, in a field no reader of a slide ever sees and therefore
 * nobody checks.
 *
 * It was not caught by review: it is invisible in a diff, because a .pptx is a
 * zip and `git` shows it as `Bin 42303 -> 54876 bytes`. It arrived with the
 * SmartArt that was authored in desktop PowerPoint and shipped in the template,
 * and it had been sitting in `modern-chart.pptx` since that deck was added.
 *
 * So the check is here rather than in either deck's own test file. One rule
 * over every committed package cannot drift the way two copies of it would,
 * and it covers the deck somebody adds next — which is the case that matters,
 * since the next deck will also be authored by a person.
 */
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
// @ts-expect-error — a plain .mjs tool with no types. The field list lives THERE
// so the remedy and the check cannot read different ones.
import { namesIn } from "../scripts/scrub-docprops.mjs";

/** The parts and fields Office writes a person or an organisation into. */
const FIELDS: [string, string[]][] = [
  ["docProps/core.xml", ["dc:creator", "cp:lastModifiedBy"]],
  ["docProps/app.xml", ["Manager", "Company"]],
];

/**
 * Every .pptx the repository tracks, asked of git rather than globbed.
 *
 * A glob over the working tree would also sweep a deck somebody left lying in
 * the checkout — a merged round's output, a deck a colleague sent — and fail
 * the suite for a file that is not going anywhere. What is committed is what is
 * published, and that is the whole of what this rule is about.
 */
function committedDecks(): string[] {
  return execFileSync("git", ["ls-files", "-z", "*.pptx", "*.potx"], { encoding: "utf8" }).split("\0").filter(Boolean);
}

describe("a deck this repository publishes", () => {
  it("has decks to check at all", () => {
    // Without this the rule below passes forever on an empty list, which is the
    // shape of vacuous guard this project has caught in itself more than once.
    expect(committedDecks().length).toBeGreaterThan(0);
  });

  it("would notice a name if one were there", () => {
    /**
     * The rule below asserts an EMPTY list, and an empty list is what a
     * matcher that has stopped matching also returns. Replacing `namesIn`'s
     * body with `return []` left all of this green — so the one guard standing
     * between a public repository and somebody's name in a file nobody opens
     * was proving nothing about the matcher at all.
     *
     * The file list is already anchored above, which is the same worry caught
     * one level out and is exactly why this gap is easy to miss: a guard can
     * be defended against vacuity in one dimension and open in the other.
     *
     * Anchored on the shape Office actually writes, and on the two ways a
     * field is legitimately empty — a self-closing tag and an empty pair —
     * because reporting those would fail every clean deck in the repo.
     */
    const core = (creator: string) =>
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties>${creator}<cp:revision>2</cp:revision></cp:coreProperties>`;
    const tags = ["dc:creator", "cp:lastModifiedBy"];

    expect(namesIn(core("<dc:creator>Ada Lovelace</dc:creator>"), tags)).toEqual(['dc:creator="Ada Lovelace"']);
    expect(namesIn(core("<cp:lastModifiedBy>Ada Lovelace</cp:lastModifiedBy>"), tags)).toEqual([
      'cp:lastModifiedBy="Ada Lovelace"',
    ]);
    // An attribute on the tag must not hide the name.
    expect(namesIn(core('<dc:creator xml:lang="en">Ada</dc:creator>'), tags)).toEqual(['dc:creator="Ada"']);
    // And the empty spellings a scrubbed deck carries are not names.
    expect(namesIn(core("<dc:creator/>"), tags)).toEqual([]);
    expect(namesIn(core("<dc:creator></dc:creator>"), tags)).toEqual([]);
    expect(namesIn(core("<dc:creator>   </dc:creator>"), tags)).toEqual([]);
  });

  it.each(committedDecks())("names nobody in its properties: %s", async (path) => {
    const zip = await JSZip.loadAsync(readFileSync(path));
    const held: string[] = [];
    for (const [part, tags] of FIELDS) {
      const file = zip.file(part);
      if (!file) continue;
      held.push(...(namesIn(await file.async("string"), tags) as string[]).map((h) => `${part} ${h}`));
    }
    expect(held, `${path} carries ${held.join(", ")} — run \`node scripts/scrub-docprops.mjs ${path}\``).toEqual([]);
  });
});
