import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

/**
 * The Marketplace listing, checked against Microsoft's rules for it.
 *
 * `docs/listing/LISTING.md` exists so the listing fields are reviewable and
 * diffable rather than retyped into a web form from memory. Nothing read it
 * until now, which meant the two numbers it quotes and the tag list it claims
 * to obey were prose: true when written and free to rot afterwards.
 *
 * The limits below are Microsoft's, from
 * https://learn.microsoft.com/en-us/partner-center/marketplace-offers/create-effective-office-store-listings
 * and https://learn.microsoft.com/en-us/partner-center/marketplace-offers/supported-html-tags
 * They are hard-coded here because Partner Center enforces them at submission
 * and nothing in this repository can discover them at run time. If Microsoft
 * changes one, this file is the place it gets changed.
 */

const listing = readFileSync("docs/listing/LISTING.md", "utf8");

/** Maximum length of the Summary field, in characters. */
const SUMMARY_MAX = 100;

/**
 * Maximum length of the Description field, in characters.
 *
 * Microsoft's own page states 10,000 in its table and 4,000 in the prose above
 * it. This asserts the smaller of the two: a description that satisfies 4,000
 * satisfies both readings, and one that satisfies only 10,000 is a coin toss
 * decided by whichever number Partner Center's form actually implements.
 */
const DESCRIPTION_MAX = 4_000;

/** The recommended word range for the Description, from the same page. */
const WORDS_MIN = 300;
const WORDS_MAX = 500;

/**
 * Every HTML tag Microsoft documents as supported in an offer description.
 *
 * Note what is missing: no `code`, so a placeholder cannot be marked up as one,
 * and no `a`, so a link cannot be embedded. The field's own help in Partner
 * Center says only "simple HTML tags" without naming any, which is how a
 * markdown draft gets written for it.
 */
const SUPPORTED_TAGS = new Set(["b", "i", "br", "p", "ul", "ol", "li", "h1", "h2", "h3", "h4", "h5", "h6"]);

/**
 * The fenced block under a heading.
 *
 * Reads the listing doc the way a person copying a field out of it does: find
 * the section, take the first fenced block under it. A section that loses its
 * block fails rather than passing vacuously, because there is nothing to
 * measure and the test says so.
 */
function block(heading: string, lang = ""): string {
  const from = listing.indexOf(`## ${heading}`);
  expect(from, `LISTING.md has no "## ${heading}" section`).toBeGreaterThan(-1);
  const rest = listing.slice(from);
  const fence = rest.match(new RegExp("```" + lang + "\\n([\\s\\S]*?)```"));
  expect(fence?.[1], `the ${heading} section has no \`\`\`${lang} block to check`).toBeTruthy();
  return (fence?.[1] ?? "").trim();
}

/** What is left of an HTML description once the markup is taken out. */
function textOf(html: string): string {
  return html
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

describe("the Summary field", () => {
  const summary = block("Summary");

  it("is one line, because the field refuses line breaks", () => {
    expect(summary).not.toContain("\n");
  });

  it(`is at most ${SUMMARY_MAX} characters`, () => {
    expect(summary.length).toBeLessThanOrEqual(SUMMARY_MAX);
  });

  it("is a sentence, because the field asks for one", () => {
    expect(summary).toMatch(/[.!?]$/);
  });
});

describe("the Description field", () => {
  const html = block("Description", "html");
  const text = textOf(html);

  it("uses only the tags Microsoft documents as supported", () => {
    const used = [...html.matchAll(/<\/?([a-z0-9]+)[^>]*>/gi)].map((m) => (m[1] ?? "").toLowerCase());
    const unsupported = [...new Set(used)].filter((t) => !SUPPORTED_TAGS.has(t));
    expect(unsupported).toEqual([]);
  });

  it("is HTML rather than markdown, which the field cannot render", () => {
    // The three that would ship visibly: bold, a bullet at the start of a line,
    // and the backticks an earlier draft put around {{Revenue|number:0}}.
    expect(html).not.toMatch(/\*\*/);
    expect(html).not.toMatch(/^- /m);
    expect(html).not.toContain("`");
  });

  it(`is at most ${DESCRIPTION_MAX} characters, tags included`, () => {
    expect(html.length).toBeLessThanOrEqual(DESCRIPTION_MAX);
  });

  it(`reads as ${WORDS_MIN} to ${WORDS_MAX} words once the tags are taken out`, () => {
    const words = text.split(" ").length;
    expect(words).toBeGreaterThanOrEqual(WORDS_MIN);
    expect(words).toBeLessThanOrEqual(WORDS_MAX);
  });

  /**
   * Microsoft: the listing description "should match the description in your
   * manifest as closely as possible". Closely is not a number, so this asserts
   * the one part that can be: the manifest's first sentence opens the listing.
   * It also ties the two files together, which is the point — the manifest is
   * generated and the listing is written by hand, so they drift silently.
   */
  it("opens on the manifest's own description", () => {
    const manifest = readFileSync("manifest-prod.xml", "utf8");
    const described = manifest.match(/<Description DefaultValue="([^"]+)"/)?.[1] ?? "";
    expect(described, "manifest-prod.xml has no Description").toBeTruthy();
    const opening = described.split(". ")[0] ?? "";
    expect(opening.length, "the manifest description has no first sentence").toBeGreaterThan(10);
    expect(text.startsWith(opening)).toBe(true);
  });
});
