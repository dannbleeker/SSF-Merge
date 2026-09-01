import { describe, expect, it } from "vitest";
import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import * as prose from "../scripts/without-prose.mjs";

const { withoutHashComments, withoutTsComments, withoutTsProse, withoutXmlComments } = prose;

/**
 * The cases that caught the guards.
 *
 * Each `it` below is a real file shape from this repo's own history, not an
 * invented one. A stripper that stops handling its own case puts the guard that
 * depends on it back where it started.
 */
const xml = withoutXmlComments as (s: string) => string;
const hash = withoutHashComments as (s: string) => string;
const ts = withoutTsProse as (s: string) => string;

describe("XML prose", () => {
  it("removes a comment that names the very element the guard forbids", () => {
    // The manifest explains why it has no <Requirements>, in a comment
    // containing the word — so the generator refused to write a correct file.
    const manifest = `<OfficeApp>
  <!-- No <Requirements>. The floor is checked at runtime by checkFloor. -->
  <Hosts><Host Name="Presentation" /></Hosts>
</OfficeApp>`;
    expect(manifest).toContain("<Requirements>");
    expect(xml(manifest)).not.toContain("<Requirements>");
    expect(xml(manifest), "the markup survives").toContain('<Host Name="Presentation" />');
  });

  it("removes a commented-out URL, which would otherwise count as one", () => {
    expect(xml('<a><!-- https://localhost:3000 --><b href="https://example.test" /></a>')).not.toContain("localhost");
  });

  it("leaves a file with no comments exactly as it was", () => {
    const clean = "<OfficeApp><Id>x</Id></OfficeApp>";
    expect(xml(clean)).toBe(clean);
  });
});

describe("hash prose", () => {
  it("removes a YAML comment naming the step the guard looks for", () => {
    // The release workflow's header explains why the tag is created by
    // `gh release create`, so the ORDER check read the comment's position.
    const workflow = `name: Release
# The tag is created by gh release create, because the proxy rejects a push.
jobs:
  release:
    steps:
      - run: node scripts/check-release.mjs
      - run: gh release create "v1.0.0"`;
    expect(workflow.indexOf("gh release create")).toBeLessThan(workflow.indexOf("check-release"));
    const stripped = hash(workflow);
    expect(stripped.indexOf("gh release create")).toBeGreaterThan(stripped.indexOf("check-release"));
  });

  it("keeps a hash INSIDE a value, which is not a comment", () => {
    // A slide id on this host looks like 256#3561048925, and a colour is #fff.
    expect(hash('  colour: "#00254C"')).toContain("#00254C");
  });
});

describe("TypeScript prose", () => {
  it("removes a docblock that names what the guard forbids", () => {
    const file = `/**
 * The seam: this file imports nothing from Office.js, deliberately.
 */
export const x = 1;`;
    expect(file).toContain("Office.js");
    expect(ts(file)).not.toContain("Office.js");
    expect(ts(file)).toContain("export const x = 1;");
  });

  it("removes a line comment and a continuation line", () => {
    expect(ts("// PowerPoint.run is avoided here\nconst y = 2;")).not.toContain("PowerPoint.");
  });

  it("removes STRING LITERALS, because a verdict naming an issue is prose", () => {
    // A sentence about office-js#6105 is not a dependency on it, and the first
    // version of the src/host guard failed on a file with no imports at all.
    const file = 'const detail = "office-js#6105 does not reproduce on this host";';
    expect(ts(file)).not.toContain("office-js");
  });

  it("does not eat code that merely follows a comment", () => {
    const file = "const a = 1; // note\nconst b = 2;";
    expect(ts(file)).toContain("const b = 2;");
  });
});

describe("a comment at the END of a line of code", () => {
  /**
   * The position the stripper did not cover, found on 2026-09-01. Block
   * comments went, and lines that BEGIN with `//` or `*` went, and a comment
   * written after code stayed — so `test/architecture.test.ts` went red on a
   * file whose only offence was the sentence `// the engine never calls
   * PowerPoint.run`. That is precisely the failure this module was written to
   * end, in the one place it was not looking.
   *
   * It belongs to `withoutTsProse` and not to `withoutTsComments`, which keeps
   * literals deliberately: a URL in a string is full of `//`, and there is no
   * way to tell one from a comment without first removing the strings.
   */
  it("goes, so a sentence after code cannot fail a guard", () => {
    expect(withoutTsProse("const x = 1; // never calls PowerPoint.run")).not.toMatch(/PowerPoint/);
  });

  it("does not take a URL inside a string with it", () => {
    // The reason this cannot live in `withoutTsComments`. The literal is
    // blanked first, so by the time comments are cut there is no URL left to
    // mistake for one — and the code around it survives.
    const stripped = withoutTsProse('const u = "https://example.com/a";\nconst after = 2;');
    expect(stripped).toContain("const after = 2");
    expect(stripped).toContain("const u =");
  });

  it("does not truncate a regex that ends in an escaped slash", () => {
    // `/https:\/\//` puts two slashes together — the escaped one and the
    // terminator — and a regex literal is not blanked. The first version of the
    // trailing-comment rule cut the file from there to the end of the line, and
    // its comment claimed that could not happen.
    const stripped = withoutTsProse("const re = /https:\\/\\//;\nconst after = 2;");
    expect(stripped).toContain("const after = 2");
    expect(stripped).toContain("/;");
  });
});

describe("withoutTsComments", () => {
  it("drops the comments and KEEPS the strings", () => {
    // The split that made this a separate export. `read-answers.mjs` reads half
    // its sheet inside template literals, so a guard that checked it through
    // `withoutTsProse` had those reads stripped out from under it and reported
    // a correct file as broken — the same false-red this module exists to stop,
    // in a fourth syntax.
    const src = [
      "/* sheet.deckRead is explained here */",
      "// and here: sheet.deckRead",
      "console.log(`${sheet.deckRead}`);",
    ].join("\n");
    const out = withoutTsComments(src);
    expect(out).toContain("sheet.deckRead}`");
    expect(out).not.toContain("explained here");
    expect(out).not.toContain("and here");
  });
});

describe("the limit of a stripper that is not a parser", () => {
  /**
   * These are deliberately NOT parsers — the module says so, and each is the
   * smallest thing that makes its own guard honest. That trade has a boundary,
   * and this is where it is.
   *
   * `withoutTsComments` finds block comments with a regex, so a `/*` inside a
   * STRING opens one and the next `*` followed by a slash closes it. Everything
   * between them leaves the file the guard is reading. The guard then reports
   * that the file does not do a thing it plainly does — which is the shape
   * three separate guards in this repo have already failed in, and the one
   * failure direction that looks like success.
   *
   * Pinned rather than fixed. Fixing it means tokenising, which is the parser
   * the module refuses to be, and no file in this repo trips it — the sweep
   * below is what says so, and it is what will fail on the day one does.
   */
  const comments = withoutTsComments as (s: string) => string;

  it("loses the code between two strings that open and close a comment", () => {
    const src = ['const OPEN = "/*";', "const kept = realCall();", 'const CLOSE = "*/";'].join("\n");
    expect(comments(src), "the limitation moved — check the sweep below still guards it").not.toContain("realCall");
  });

  it("is untroubled by an opener with no closer", () => {
    // The reassuring half, and the reason this has never bitten: it takes BOTH.
    const src = ['const u = "https://example.com/*";', "const kept = realCall();"].join("\n");
    expect(comments(src)).toContain("realCall");
  });
});

describe("no file in this repo trips that limit", () => {
  /**
   * The sweep that makes the pin above safe to leave. Every `export` a file
   * declares has to survive `withoutTsComments`, which keeps literals and
   * removes only prose.
   *
   * An export named inside a COMMENT is expected to disappear with it — this
   * repo's doc comments quote declarations constantly — so those are allowed by
   * checking whether the line it was found on is a comment line. Without that
   * the sweep reports `scripts/sibling-watch.mjs`, whose docstring quotes
   * `export const NAME`, and a sweep that cries wolf is one somebody widens
   * until it is quiet.
   */
  const comments = withoutTsComments as (s: string) => string;
  const DECLARATION = /export (?:async )?(?:function|const|class|interface|type) ([A-Za-z0-9_]+)/g;

  function sourceFiles(dir: string): string[] {
    return readdirSync(dir, { withFileTypes: true }).flatMap((e) =>
      e.isDirectory() ? sourceFiles(join(dir, e.name)) : /\.(ts|mjs)$/.test(e.name) ? [join(dir, e.name)] : [],
    );
  }

  it("keeps every declaration a file makes outside a comment", () => {
    const files = ["src", "scripts", "test"].flatMap(sourceFiles);
    expect(files.length, "the sweep stopped finding any source").toBeGreaterThan(40);

    const lost: string[] = [];
    for (const file of files) {
      const src = readFileSync(file, "utf8");
      const kept = new Set([...comments(src).matchAll(DECLARATION)].map((m) => m[1] as string));
      for (const line of src.split("\n")) {
        // A line of prose that QUOTES a declaration is not a declaration.
        if (/^\s*(\*|\/\/)/.test(line)) continue;
        for (const m of line.matchAll(DECLARATION)) {
          if (!kept.has(m[1] as string)) lost.push(`${file}: ${m[1] as string}`);
        }
      }
    }
    expect(lost, "the stripper ate a declaration, so a guard is reading a file that is not there").toEqual([]);
  });
});
