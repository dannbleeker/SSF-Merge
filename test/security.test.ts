/**
 * The security sweep, as tests.
 *
 * `SECURITY.md` makes claims about how this add-in handles what arrives from
 * outside — a pasted table and a .pptx, both of which a user can be sent. A
 * claim in a document is a claim nobody re-checks, so each one that can be
 * executed lives here instead, and the document points at this file.
 *
 * Swept 2026-08-29. What it found is written up in SECURITY.md.
 */
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { Pkg, resolveTarget } from "../src/core/pptx/pkg.js";
import { prepareBlock } from "../src/core/merge/prepare.js";
import { buildPlan } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import { toRecordSet } from "../src/core/data/recordset.js";
import { A_NS, elements, parseXml } from "../src/core/pptx/xml.js";
import { makeDeck } from "./fixtures/deck.js";

/** Merge one row through the ordinary path and hand back the package. */
async function mergeRows(paragraphs: string[][], table: string[][]) {
  const pkg = await Pkg.open(await makeDeck([{ paragraphs }]));
  const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "sec");
  if (!prepared.ok) throw new Error(`the fixture was refused: ${prepared.why}`);
  const records = toRecordSet(table);
  const plan = buildPlan(prepared.block, records, { runId: "sec" });
  await runPlan(pkg, plan, records);
  return JSZip.loadAsync(await pkg.toBytes());
}

/**
 * The text a PARSER sees, not the bytes.
 *
 * Read with a regex over the raw XML this answers `&lt;b&gt;`, which is the
 * escaping doing its job and looks like a failed round trip. The question is
 * what PowerPoint reads back, so parse it the way PowerPoint would.
 */
const textOf = (xml: string) =>
  elements(parseXml(xml), A_NS, "t")
    .map((n) => n.textContent ?? "")
    .join("");

describe("a pasted value is data, never markup", () => {
  /**
   * The value goes in through `textContent`, and @xmldom escapes on the way
   * out. That is the whole mechanism — there is no escape() to call and no
   * caller who could forget to.
   *
   * Checked as a ROUND TRIP rather than by looking for entities, because the
   * failure that actually happened on this project was the opposite one: a
   * value escaped twice, so `Ben & Co` merged as `Ben &amp;amp; Co` and reached
   * the slide with the entity showing.
   */
  it("survives a merge exactly once, however hostile", async () => {
    const nasty = [
      "<b>bold</b>",
      "a & b",
      "]]>break",
      `"quoted" <!-- comment --> 'apostrophe'`,
      "<?xml version='1.0'?><!DOCTYPE x>",
      "&amp; already an entity",
    ];
    for (const value of nasty) {
      const zip = await mergeRows([["{{V}}"]], [["V"], [value]]);
      const merged = await zip.file("ppt/slides/slide2.xml")!.async("string");
      expect(textOf(merged), `round trip of ${JSON.stringify(value)}`).toBe(value);
    }
  });

  it("leaves every part of the package parseable", async () => {
    // The harm from a broken escape is not a script running — there is nowhere
    // for one to run. It is a deck PowerPoint refuses to open, which is the
    // same harm as a wrong merge and looks like the add-in's fault either way.
    const zip = await mergeRows([["{{V}}"]], [["V"], ["</a:t></a:r></a:p><p:evil/>"]]);
    for (const name of Object.keys(zip.files).filter((n) => n.endsWith(".xml"))) {
      const xml = await zip.file(name)!.async("string");
      expect(() => parseXml(xml), name).not.toThrow();
      expect(xml, `${name} carries an injected element`).not.toContain("<p:evil/>");
    }
  });
});

describe("a column name is data, never a property", () => {
  it("keeps __proto__ and constructor as ordinary columns", () => {
    // `Object.create(null)` for the row, and every read goes through
    // hasOwnProperty. Both, deliberately: either alone would be enough until
    // somebody changes the other.
    const set = toRecordSet([
      ["__proto__", "constructor", "toString"],
      ["P", "C", "T"],
    ]);
    expect(set.columns.map((c) => c.name)).toEqual(["__proto__", "constructor", "toString"]);
    const row = set.rows[0]!;
    expect(Object.getPrototypeOf(row)).toBeNull();
    expect(Object.keys(row).sort()).toEqual(["__proto__", "constructor", "toString"]);
    expect(row["__proto__"]).toBe("P");
  });

  it("pollutes nothing", () => {
    toRecordSet([["__proto__"], ['{"polluted":true}']]);
    expect(({} as Record<string, unknown>)["polluted"]).toBeUndefined();
    expect(Object.prototype).not.toHaveProperty("polluted");
  });

  it("merges a value under such a column onto the slide", async () => {
    const zip = await mergeRows([["{{__proto__}}"]], [["__proto__"], ["landed"]]);
    expect(textOf(await zip.file("ppt/slides/slide2.xml")!.async("string"))).toBe("landed");
  });
});

describe("a relationship target comes out of the deck", () => {
  /**
   * `resolveTarget` honours a leading `/` and any number of `..`, which is what
   * the format says to do — so it can and does answer a path outside the part's
   * own directory. That is fine for READING, where a path that names nothing
   * simply is not found. It matters where the answer is used to DELETE.
   */
  it("can name a part outside the slide's directory", () => {
    const owner = "ppt/slides/slide1.xml";
    expect(resolveTarget(owner, "/[Content_Types].xml")).toBe("[Content_Types].xml");
    expect(resolveTarget(owner, "../../[Content_Types].xml")).toBe("[Content_Types].xml");
    // `..` past the root is clamped by `pop()` on an empty array rather than
    // escaping into something above it: the answer stays a package-relative
    // name, so nothing here can reach a real filesystem.
    expect(resolveTarget(owner, "../../../../../../etc/passwd")).toBe("etc/passwd");
  });

  it("does not let a crafted notes target delete a part the deck needs", async () => {
    // A deck arrives from anywhere. Removing a slide collects that slide's
    // notes page and comments, and before this was guarded a target of
    // `/[Content_Types].xml` had it delete the one part a presentation cannot
    // open without — turning a merge of somebody else's deck into a file that
    // will not open.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["Cover"]] }, { paragraphs: [["Second"]] }]));
    const relsPath = Pkg.relsPathFor("ppt/slides/slide1.xml");
    const rels =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
      `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="/[Content_Types].xml"/>` +
      `</Relationships>`;
    pkg.setText(relsPath, rels);

    expect(pkg.has("[Content_Types].xml"), "the fixture should start with one").toBe(true);
    await pkg.removeSlide("ppt/slides/slide1.xml");
    expect(pkg.has("[Content_Types].xml"), "a crafted target deleted it").toBe(true);
  });
});

describe("a chart's own relationships come out of the deck too", () => {
  /**
   * The same class of hole as the one above, one level down — found only by
   * looking for siblings after fixing the first, which is the step that gets
   * skipped.
   *
   * Removing a slide sweeps the parts its charts and diagrams own. The parent
   * of that list was held to an allowlist; the CHILD was not. So a chart whose
   * relationships named `/ppt/presentation.xml` had that part counted as
   * something the chart owned, nothing else in the package referred to it — its
   * only referrer is the root `_rels/.rels`, which the referrer scan does not
   * read — and the sweep took it.
   *
   * `/ppt/presentation.xml` was the worse one: deleted SILENTLY, the merge
   * finished, and the output cannot open. `/[Content_Types].xml` was deleted and
   * then threw.
   */
  async function deckWhoseChartClaims(victim: string) {
    const pkg = await Pkg.open(
      await makeDeck([
        { paragraphs: [["Cover"]], chart: { title: "T", workbook: ["a"] } },
        { paragraphs: [["Second"]] },
      ]),
    );
    pkg.setText(
      "ppt/charts/_rels/chart1.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="${victim}"/>` +
        `</Relationships>`,
    );
    return pkg;
  }

  for (const victim of ["/ppt/presentation.xml", "/[Content_Types].xml", "/ppt/slides/slide2.xml"]) {
    it(`keeps ${victim} when a chart claims to own it`, async () => {
      const pkg = await deckWhoseChartClaims(victim);
      const key = victim.slice(1);
      expect(pkg.has(key), "the fixture should start with it").toBe(true);
      await pkg.removeSlide("ppt/slides/slide1.xml");
      expect(pkg.has(key), `a crafted chart relationship swept ${key}`).toBe(true);
    });
  }

  it("still sweeps what the slide's chart really does own", async () => {
    // The guard must not become "never sweep anything": a template slide going
    // out on the whole-deck route has to take its chart with it, or the output
    // ships parts nothing points at.
    const pkg = await Pkg.open(
      await makeDeck([
        { paragraphs: [["Cover"]], chart: { title: "T", workbook: ["a"] } },
        { paragraphs: [["Second"]] },
      ]),
    );
    expect(pkg.has("ppt/charts/chart1.xml")).toBe(true);
    await pkg.removeSlide("ppt/slides/slide1.xml");
    expect(pkg.has("ppt/charts/chart1.xml"), "the chart should have gone with its slide").toBe(false);
  });
});
