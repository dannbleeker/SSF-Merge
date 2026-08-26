import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { creationIdOf } from "../src/core/pptx/clone.js";
import { Pkg } from "../src/core/pptx/pkg.js";
import { TAG_RUN, readSlideTags } from "../src/core/pptx/tags.js";

/**
 * These read the ARTIFACT, not a recipe for it.
 *
 * The probe's whole value is that its two arms differ in exactly one way, and a
 * builder that quietly produced two identical decks would give an answer sheet
 * that looks complete and says nothing. Reading the generated snippet means
 * there is no second copy of the recipe to drift from the first, and it fails
 * when the snippet has not been regenerated.
 */
const snippet = readFileSync("probe/probe-snippet.ts", "utf8");

function deckFromSnippet(name: string): string {
  // Tolerant of whitespace on purpose: a formatter that wrapped the
  // assignment once broke this silently, and a guard that stops matching is a
  // guard that passes forever.
  const match = new RegExp(`const ${name}\\s*=\\s*"([A-Za-z0-9+/=]+)"`).exec(snippet);
  if (!match?.[1]) throw new Error(`the snippet has no ${name}`);
  return match[1];
}

describe("the probe's fixture decks", () => {
  it("both hold two slides, so the arms differ only in the creation ids", async () => {
    for (const name of ["FRESH_DECK", "COLLISION_DECK"]) {
      const pkg = await Pkg.open(deckFromSnippet(name));
      expect(await pkg.slidePaths(), name).toHaveLength(2);
    }
  });

  it("gives the fresh arm two DIFFERENT creation ids", async () => {
    const pkg = await Pkg.open(deckFromSnippet("FRESH_DECK"));
    const ids = await Promise.all((await pkg.slidePaths()).map((p) => creationIdOf(pkg, p)));
    expect(new Set(ids).size).toBe(2);
  });

  it("gives the collision arm ONE creation id shared by both slides", async () => {
    // The arm that tests office-js#6105. If this ever produced two distinct ids
    // the probe would report a clean insert and prove nothing, which is the
    // vacuous measurement this project keeps having to catch.
    const pkg = await Pkg.open(deckFromSnippet("COLLISION_DECK"));
    const ids = await Promise.all((await pkg.slidePaths()).map((p) => creationIdOf(pkg, p)));
    expect(new Set(ids).size).toBe(1);
    expect(ids[0]).toBeDefined();
  });

  it("tags the fresh arm's slides in the package, which is what question two reads back", async () => {
    const pkg = await Pkg.open(deckFromSnippet("FRESH_DECK"));
    for (const path of await pkg.slidePaths()) {
      expect((await readSlideTags(pkg, path)).get(TAG_RUN), path).toBe("probe-run");
    }
  });
});

describe("the probe snippet", () => {
  it("bounds every host call, because a stall here is death rather than slowness", () => {
    expect(snippet).toContain("withTimeout");
    expect(snippet).not.toMatch(/await PowerPoint\.run\([\s\S]{0,40}\)\s*;\s*\/\/ unbounded/);
  });

  it("cleans up by position and never by id", () => {
    // A slide this run added does not round-trip through slides.getItem(id) on
    // the web, and a sibling project's by-id clean-up reported 45 successful
    // deletes having removed nothing.
    expect(snippet).toContain("getItemAt(i).delete()");
    expect(snippet).not.toContain("getItem(id).delete");
  });

  it("passes a targetSlideId, so the probe's slides do not land in front of the user's", () => {
    expect(snippet).toContain("targetSlideId");
  });

  it("carries a floor that stops the sweep reaching the user's own slides", () => {
    expect(snippet).toContain("from < deckAtStart");
  });
});

describe("the probe's fixture package", () => {
  it("carries a theme with all three of its required children", async () => {
    // CT_BaseStyles requires clrScheme, fontScheme and fmtScheme, every one
    // mandatory. This part was `<a:themeElements/>` when the first real sheet
    // came back InvalidArgument from every insert, and KeepSourceFormatting is
    // exactly the path that has to import the source theme. Checked against a
    // deck PowerPoint itself accepts, which carries all three.
    const pkg = await Pkg.open(deckFromSnippet("FRESH_DECK"));
    const theme = await pkg.text("ppt/theme/theme1.xml");
    for (const child of ["clrScheme", "fontScheme", "fmtScheme"]) {
      expect(theme).toContain(`<a:${child} `);
    }
  });

  it("names every part it declares a content type for", async () => {
    // A content-type Override for a part that is not in the zip is a package
    // no reader will open, and it is the easy mistake when adding parts.
    const pkg = await Pkg.open(deckFromSnippet("FRESH_DECK"));
    const types = await pkg.text("[Content_Types].xml");
    const parts = [...types.matchAll(/<Override PartName="\/([^"]+)"/g)].map((m) => m[1] ?? "");
    expect(parts.length).toBeGreaterThan(4);
    for (const part of parts) {
      await expect(pkg.text(part)).resolves.toBeTypeOf("string");
    }
  });
});

describe("the generated snippet", () => {
  it("typechecks against the real Office.js types", () => {
    // The snippet is outside tsconfig's include and is pasted into an editor
    // that will run it before anyone reads it, so nothing else here would catch
    // a misspelled option key or a call that does not exist. Proven non-vacuous
    // by adding an unknown key to InsertSlideOptions: tsc names it.
    expect(() =>
      execFileSync(
        process.execPath,
        [
          "./node_modules/typescript/bin/tsc",
          "--noEmit",
          "--lib",
          "es2020,dom",
          "--types",
          "office-js",
          "--skipLibCheck",
          "probe/probe-snippet.ts",
        ],
        { encoding: "utf8", stdio: "pipe" },
      ),
    ).not.toThrow();
  }, 60000);

  it("asks its control arm BEFORE it adds anything of its own", () => {
    // The control inserts the presentation's own bytes. Run later it would be
    // inserting the probe's slides back too, which grows the deck by whatever
    // the earlier arms happened to land and makes its own reading unreadable.
    const control = snippet.indexOf("answers.insertSelf");
    const first = snippet.indexOf("answers.insertFresh");
    expect(control).toBeGreaterThan(-1);
    expect(control).toBeLessThan(first);
  });

  it("names the call that threw rather than the whole probe", () => {
    // Four calls shared one catch, so the first real sheet said
    // "InvalidArgument" about a statement nobody could identify.
    expect(snippet).toContain("failedAt: step");
    expect(snippet).toMatch(/step = "adding a text box"/);
  });
});
