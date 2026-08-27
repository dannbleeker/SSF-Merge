import { describe, expect, it } from "vitest";
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
