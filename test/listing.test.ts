import { existsSync, readFileSync, readdirSync } from "node:fs";
import { sep } from "node:path";
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

  it("says 'mail merge' for as long as the product name does not", () => {
    // The name is staying `SSF Merge`, which a stranger cannot expand and which
    // does not say merge what. That decision is only affordable because the
    // summary spends its opening on the category instead of the outcome, so the
    // two are checked together rather than left agreeing in prose.
    //
    // It runs both ways on purpose. A rename to something carrying the words
    // would make this line redundant, and the shorter summary already drafted
    // in LISTING.md becomes the better one — which is a thing to be told about
    // rather than left paying for a cost that has gone away.
    const manifest = readFileSync("manifest-prod.xml", "utf8");
    const name = /<DisplayName DefaultValue="([^"]+)"/.exec(manifest)?.[1];
    expect(name, "manifest-prod.xml has no DisplayName").toBeTruthy();

    const nameCarriesIt = /mail\s*merge/i.test(name ?? "");
    const summaryCarriesIt = /mail\s*merge/i.test(summary);
    if (nameCarriesIt) {
      expect(
        summaryCarriesIt,
        `"${name}" already says mail merge, so the summary need not spend its opening on it`,
      ).toBe(false);
    } else {
      expect(summaryCarriesIt, `nothing says "mail merge": not "${name}", not the summary`).toBe(true);
    }
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

describe("the Marketplace icon", () => {
  /** Partner Center's own limits on the listing page's icon upload. */
  const ICON_PX = 300;
  const ICON_MAX_BYTES = 512 * 1024;
  const file = `docs/listing/marketplace-icon-${ICON_PX}.png`;

  it(`is a whole PNG, ${ICON_PX} by ${ICON_PX}`, () => {
    // A truncated or zero-byte file still satisfies existsSync, and Partner
    // Center rejecting the upload is a slower way to learn that than this.
    //
    // The IEND check is here because the header check alone did NOT catch a
    // truncation: cutting this file to its first 900 bytes left the signature
    // and the IHDR intact, so width and height still read 300 and only the
    // byte-identity test below noticed. A PNG ends with a twelve-byte IEND
    // chunk, and a file that stops early does not have one.
    const bytes = readFileSync(file);
    expect(bytes.subarray(0, 8).toString("hex"), "PNG signature").toBe("89504e470d0a1a0a");
    expect(bytes.readUInt32BE(16), "width").toBe(ICON_PX);
    expect(bytes.readUInt32BE(20), "height").toBe(ICON_PX);
    expect(bytes.subarray(-8).toString("hex"), "IEND chunk").toBe("49454e44ae426082");
  });

  it("is under the 512 KB the field accepts", () => {
    expect(readFileSync(file).length).toBeLessThanOrEqual(ICON_MAX_BYTES);
  });

  it("is byte-identical to what scripts/build-icons.mjs draws", async () => {
    // The same reason the manifest's icons are pinned: a committed binary is
    // the one kind of file a reviewer cannot read. If this fails, either the
    // drawing changed and the PNG was not rebuilt, or it was edited by hand.
    // @ts-expect-error — plain .mjs with no types.
    const icons = await import("../scripts/build-icons.mjs");
    const drawn = icons.png(icons.MARKETPLACE, icons.supersampled(icons.MARKETPLACE)) as Buffer;
    expect(Buffer.compare(drawn, readFileSync(file))).toBe(0);
  });

  /**
   * Microsoft: "both images should be of the same logo or icon. This way, the
   * user sees the same logo in Microsoft Marketplace and when the solution is
   * displayed in Office."
   *
   * Asserted on the drawing rather than by comparing pixels across a 300 and a
   * 32, which differ legitimately: the listing icon is supersampled and the
   * ribbon icons are not. What has to hold is that both come from `markPixel`,
   * so a change to the mark cannot reach one and miss the other.
   */
  it("is the same mark the manifest's icons draw", async () => {
    // @ts-expect-error — plain .mjs with no types.
    const { markPixel, supersampled, MARKETPLACE, NAVY, ORANGE } = await import("../scripts/build-icons.mjs");
    const big = supersampled(MARKETPLACE) as (x: number, y: number) => number[];
    const small = markPixel(32) as (x: number, y: number) => number[];

    // The centre of each icon is ground in both, and the orange row sits at the
    // same fraction of the height in both.
    const orangeRowAt = (at: (x: number, y: number) => number[], size: number) => {
      for (let y = 0; y < size; y++) {
        const [r, g, b] = at(Math.round(size / 2), y);
        if (r === ORANGE[0] && g === ORANGE[1] && b === ORANGE[2]) return y / size;
      }
      return -1;
    };
    const wide = orangeRowAt(big, MARKETPLACE);
    const narrow = orangeRowAt(small, 32);
    expect(wide, "no orange row in the marketplace icon").toBeGreaterThan(0);
    expect(narrow, "no orange row in icon-32").toBeGreaterThan(0);
    expect(Math.abs(wide - narrow)).toBeLessThan(0.02);

    const [r, g, b] = big(2, Math.round(MARKETPLACE / 2));
    expect([r, g, b], "the ground is navy in both").toEqual(NAVY);
  });
});

describe("the Screenshots field", () => {
  /** Partner Center's own limits on the listing page's screenshot uploads. */
  const SHOT_W = 1366;
  const SHOT_H = 768;
  const SHOT_MAX_BYTES = 1024 * 1024;
  const SHOT_MAX_COUNT = 5;
  const dir = "docs/listing/shots";

  /** The ones meant for upload: `1-` to `5-`. Anything else is a spare. */
  const uploads = readdirSync(dir)
    .filter((f) => /^\d-.*\.png$/.test(f))
    .sort();

  it(`has at least one and no more than ${SHOT_MAX_COUNT}`, () => {
    // Partner Center takes up to five. Shooting a sixth is cheap and choosing
    // between them is not, so the extras stay in the folder under a name that
    // does not start with a digit rather than being deleted.
    expect(uploads.length).toBeGreaterThan(0);
    expect(uploads.length).toBeLessThanOrEqual(SHOT_MAX_COUNT);
  });

  it("is numbered from 1 with no gaps, because the store shows them in order", () => {
    expect(uploads.map((f) => f[0])).toEqual(uploads.map((_, i) => String(i + 1)));
  });

  it.each(uploads)(`%s is a whole PNG at exactly ${SHOT_W} by ${SHOT_H}`, (name) => {
    // Exactly, not merely close. Partner Center scales anything else, and the
    // pane's text is small enough that resampling visibly softens it — which is
    // the whole reason the capture emulates the viewport instead of cropping.
    //
    // The IEND check catches the truncation that the header check does not: a
    // file cut short keeps its signature and its IHDR, so width and height
    // still read correctly on a file no decoder will open.
    const bytes = readFileSync(`${dir}/${name}`);
    expect(bytes.subarray(0, 8).toString("hex"), "PNG signature").toBe("89504e470d0a1a0a");
    expect(bytes.readUInt32BE(16), "width").toBe(SHOT_W);
    expect(bytes.readUInt32BE(20), "height").toBe(SHOT_H);
    expect(bytes.subarray(-8).toString("hex"), "IEND chunk").toBe("49454e44ae426082");
    expect(bytes.length, "under the 1024 KB the field accepts").toBeLessThanOrEqual(SHOT_MAX_BYTES);
  });

  it("is not the old placeholder pair", () => {
    // Those two were real captures of the real product and were still wrong to
    // upload: a crop fixture that reads as a broken image, a title bar saying
    // SSF-Merge-test-template, and a photograph of a real person in the corner.
    // They were deleted rather than renamed, so this asserts they are gone.
    for (const old of ["docs/listing/01-attach-your-rows.png", "docs/listing/02-see-what-it-will-add.png"]) {
      expect(existsSync(old), `${old} is a placeholder and must never be uploaded`).toBe(false);
    }
  });
});

describe("the Search keywords field", () => {
  const keywords = block("Search keywords")
    .split("\n")
    .map((k) => k.trim())
    .filter(Boolean);

  it("has at most the three the field accepts", () => {
    expect(keywords.length).toBeGreaterThan(0);
    expect(keywords.length).toBeLessThanOrEqual(3);
  });

  /**
   * Partner Center's help on the field: "Don't add words or acronyms that are
   * already included in your product's name, summary, or description."
   *
   * Nothing Microsoft publishes covers this field, so that sentence is the whole
   * rule, and a repeated word is a wasted slot out of three. The name comes from
   * the manifest rather than from prose in the listing doc, because Microsoft
   * requires the two to match — "You specify your add-in name in two places, so
   * be sure to use the same name in both" — which means renaming the product
   * moves this guard with it and cannot leave it checking the old name.
   */
  it("spends no slot on a word the name, summary or description already has", () => {
    const manifest = readFileSync("manifest-prod.xml", "utf8");
    const name = manifest.match(/<DisplayName DefaultValue="([^"]+)"/)?.[1] ?? "";
    expect(name, "manifest-prod.xml has no DisplayName").toBeTruthy();

    const spent = new Set(
      `${name} ${block("Summary")} ${textOf(block("Description", "html"))}`.toLowerCase().match(/[a-z]+/g),
    );
    const repeated = keywords.filter((k) => spent.has(k.toLowerCase()));
    expect(repeated).toEqual([]);
  });
});

describe("the Notes for certification field", () => {
  const notes = block("Notes for certification");

  it("says up front that nothing has to be bought or signed into", () => {
    // The field's own warning is about exactly this, and it threatens "an
    // automatic rejection" for leaving it out. The assertion is on the first
    // sentence rather than anywhere in the text: a reviewer who has to hunt
    // for it has already been given a reason to doubt the rest.
    const first = notes.split("\n\n")[0];
    expect(first).toMatch(/no test account, license key, or purchase is required/i);
  });

  it("tells the reviewer to press the button the manifest actually creates", () => {
    // The label and the tab live in manifest-prod.xml. A rename there would
    // leave this note directing a reviewer to a control that is not there,
    // which is a failed certification rather than a failed test.
    const manifest = readFileSync("manifest-prod.xml", "utf8");
    const label = /<bt:String id="OpenPane\.Label" DefaultValue="([^"]+)"/.exec(manifest)?.[1];
    expect(label, "manifest-prod.xml has no OpenPane.Label").toBeTruthy();
    expect(notes).toContain(`"${label}"`);
    expect(manifest).toContain('<OfficeTab id="TabHome">');
    expect(notes).toMatch(/on the Home tab/i);
  });

  it("hands the reviewer the same data the demo deck expects", () => {
    // The one that matters. `demo/rows.txt` is tab-separated, because that is
    // what a spreadsheet paste produces; these rows are comma-separated,
    // because a cert note is retyped out of a web form where a tab is as
    // likely to move the focus as to reach the clipboard. Two copies of the
    // same table in two formats is exactly the shape that drifts, and the
    // symptom would be a reviewer pasting rows the deck no longer matches.
    const rows = readFileSync("docs/listing/demo/rows.txt", "utf8")
      .trim()
      .split("\n")
      .map((line) => line.split("\t").map((c) => c.trim()));

    for (const row of rows) {
      expect(notes, `the note must offer the demo deck's row ${JSON.stringify(row[0])}`).toContain(row.join(","));
    }
  });

  it("names a test deck that is actually in the repository", () => {
    // A download link is only as good as the file behind it. The URL is built
    // from a repository path, so this checks the path rather than the network:
    // a moved or renamed deck breaks here instead of under a reviewer.
    const url = /https:\/\/github\.com\/\S+\/raw\/main\/(\S+\.pptx)/.exec(notes)?.[1];
    expect(url, "the note offers no downloadable test deck").toBeTruthy();
    expect(existsSync(url ?? ""), `${url} is linked from the note and not in the repository`).toBe(true);
  });

  it("does not claim a network silence the source contradicts", () => {
    // The note tells a reviewer the add-in makes no network requests while it
    // runs. That is a claim about src/, so it is checked against src/ — and a
    // reviewer with the network tab open would catch it faster than a rewrite.
    expect(notes).toMatch(/makes no network requests/i);
    const sources = readdirSync("src", { recursive: true, encoding: "utf8" })
      .filter((f) => /\.tsx?$/.test(f))
      // readdirSync's recursive mode returns Windows separators here.
      .map((f) => ["src", ...String(f).split(sep)].join("/"));
    expect(sources.length, "found no sources to check the claim against").toBeGreaterThan(0);

    const callers = sources.filter((f) =>
      /\b(fetch|XMLHttpRequest|WebSocket|sendBeacon|EventSource)\s*\(/.test(readFileSync(f, "utf8")),
    );
    expect(callers, "these make network calls, so the note's claim is false").toEqual([]);
  });
});
