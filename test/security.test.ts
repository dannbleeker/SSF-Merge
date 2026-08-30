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
import { readFileSync, readdirSync } from "node:fs";
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { Pkg, resolveTarget } from "../src/core/pptx/pkg.js";
import { prepareBlock } from "../src/core/merge/prepare.js";
import { buildPlan } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import { toRecordSet } from "../src/core/data/recordset.js";
import { A_NS, elements, parseXml } from "../src/core/pptx/xml.js";
import { makeDeck } from "./fixtures/deck.js";

const ROOT_RELS_PATH = "_rels/.rels";

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

  it("counts the package's own relationships as referrers", async () => {
    /**
     * The referrer scan tested every `.rels` whose path contained `/_rels/`,
     * which is every one EXCEPT the package's own `_rels/.rels`. So a part
     * named only from the root was invisible to it and looked unreferenced.
     *
     * `ppt/presentation.xml` is the real instance, and the allowlist above now
     * keeps it out of reach — which means this needs a part the allowlist DOES
     * admit to be observable at all. A picture referenced from the root, the
     * way a thumbnail is, and claimed by the chart on the way out.
     */
    const pkg = await Pkg.open(
      await makeDeck([
        { paragraphs: [["Cover"]], chart: { title: "T", workbook: ["a"] } },
        { paragraphs: [["Second"]] },
      ]),
    );
    pkg.setBytes("ppt/media/image9.png", new Uint8Array([137, 80, 78, 71]));
    const root = await pkg.text(ROOT_RELS_PATH);
    pkg.setText(
      ROOT_RELS_PATH,
      root.replace(
        "</Relationships>",
        `<Relationship Id="rIdPic" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/thumbnail" Target="ppt/media/image9.png"/></Relationships>`,
      ),
    );
    pkg.setText(
      "ppt/charts/_rels/chart1.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image9.png"/>` +
        `</Relationships>`,
    );

    await pkg.removeSlide("ppt/slides/slide1.xml");
    expect(pkg.has("ppt/media/image9.png"), "a part referenced only from the root was swept").toBe(true);
  });

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

describe("the XML parser does not fetch or expand what a deck tells it to", () => {
  /**
   * From the sweep of 2026-08-30. A .pptx is a zip of XML and a user can be
   * sent one, so the two classic parser attacks are the first question anybody
   * asks of this codebase — and until now the answer was a property of a
   * dependency that nothing here had checked.
   *
   * `@xmldom/xmldom` refuses both. These run so that a version bump which
   * changes its mind is a red test rather than a discovery.
   */
  it("does not resolve an external entity", () => {
    const doc = parseXml('<?xml version="1.0"?>\n<!DOCTYPE r [ <!ENTITY x SYSTEM "file:///etc/passwd"> ]>\n<r>&x;</r>');
    // The reference comes through as TEXT. Not "the file was empty" — the
    // parser never went looking, which is why the literal is still there.
    expect(doc.documentElement?.textContent).toBe("&x;");
  });

  it("does not expand a nested entity bomb", () => {
    const bomb =
      '<?xml version="1.0"?>\n<!DOCTYPE r [\n' +
      '<!ENTITY a "aaaaaaaaaa">\n' +
      '<!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;">\n' +
      '<!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;">\n' +
      '<!ENTITY d "&c;&c;&c;&c;&c;&c;&c;&c;&c;&c;">\n' +
      '<!ENTITY e "&d;&d;&d;&d;&d;&d;&d;&d;&d;&d;">\n' +
      "]>\n<r>&e;</r>";
    // Four levels is 100,000 characters if expanded, and the published attack
    // has nine. Custom entities are not expanded at all, so what comes out is
    // the reference itself.
    const text = parseXml(bomb).documentElement?.textContent ?? "";
    expect(text.length).toBeLessThan(100);
    expect(text).toBe("&e;");
  });
});

describe("the claims on the front of SECURITY.md are executable", () => {
  /**
   * That page says the add-in makes no network calls and never writes markup.
   * Both are properties of the SOURCE, so both can be read off it — and a
   * security page whose claims nothing re-checks is the failure the page's own
   * preamble warns about.
   */
  const sources = (): string[] => {
    const out: string[] = [];
    const walk = (dir: string): void => {
      for (const entry of readdirSync(dir, { withFileTypes: true })) {
        const path = `${dir}/${entry.name}`;
        if (entry.isDirectory()) walk(path);
        else if (/\.(ts|html)$/.test(entry.name)) out.push(path);
      }
    };
    walk("src");
    return out;
  };

  it("reads the source at all", () => {
    // The vacuity guard this suite asks for twice elsewhere: an empty file list
    // would satisfy every assertion below forever.
    expect(sources().length).toBeGreaterThan(20);
  });

  it("makes no network call", () => {
    for (const path of sources()) {
      const src = readFileSync(path, "utf8");
      for (const call of ["fetch(", "XMLHttpRequest", "WebSocket", "sendBeacon"]) {
        // Comments are stripped, because this repo's files explain themselves
        // and several of them name these APIs in prose.
        const code = src.replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");
        expect(code.includes(call), `${path} calls ${call}`).toBe(false);
      }
    }
  });

  it("writes text, never markup", () => {
    for (const path of sources()) {
      const code = readFileSync(path, "utf8")
        .replace(/\/\*[\s\S]*?\*\//g, "")
        .replace(/^\s*\/\/.*$/gm, "");
      for (const sink of ["outerHTML", "insertAdjacentHTML", "document.write", "eval(", "new Function"]) {
        expect(code.includes(sink), `${path} uses ${sink}`).toBe(false);
      }
      // `innerHTML` is READ once, in the page's own no-Office fallback, to ask
      // whether anything has been drawn yet. Assigning to it is the thing this
      // forbids.
      expect(/innerHTML\s*=/.test(code), `${path} assigns innerHTML`).toBe(false);
    }
  });
});
