import { describe, expect, it, vi } from "vitest";
import JSZip from "jszip";
import { cloneSlide, creationIdOf, notesPathFor, setCreationId } from "../src/core/pptx/clone.js";
import { buildPlan } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import { toRecordSet } from "../src/core/data/recordset.js";
import { prepareBlock } from "../src/core/merge/prepare.js";
import { Pkg } from "../src/core/pptx/pkg.js";
import {
  TAG_BLOCK,
  TAG_RECORD,
  TAG_RUN,
  mergeTagPart,
  tagPartXml,
  nextTagNumber,
  readSlideTags,
  writeSlideTags,
} from "../src/core/pptx/tags.js";
import { A_NS, P_NS, R_NS, child, children, element, elements, parseXml } from "../src/core/pptx/xml.js";

const P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
import { makeDeck } from "./fixtures/deck.js";

async function deck(...args: Parameters<typeof makeDeck>): Promise<Pkg> {
  return Pkg.open(await makeDeck(...args));
}

const ONE = [{ paragraphs: [["Hello {{Name}}"]], creationId: 111 }];

describe("Pkg", () => {
  it("lists slides in presentation order, not zip order", async () => {
    const pkg = await deck([{ paragraphs: [["a"]] }, { paragraphs: [["b"]] }]);
    expect(await pkg.slidePaths()).toEqual(["ppt/slides/slide1.xml", "ppt/slides/slide2.xml"]);
  });

  it("gives a new relationship the highest id plus one", async () => {
    const pkg = await deck(ONE);
    const first = await pkg.addRel("ppt/presentation.xml", "http://example/t", "x.xml");
    const second = await pkg.addRel("ppt/presentation.xml", "http://example/t", "y.xml");
    expect(Number(first.slice(3))).toBeLessThan(Number(second.slice(3)));
  });

  it("never names a part the package already holds, however big the numbers are", async () => {
    /**
     * "Highest existing number plus one" stops being that above 2^53, where
     * `max + 1 === max` — so a package holding `slide99999999999999999999.xml`
     * answered a number already in use, `copyPart` overwrote it silently, and
     * three merged slides pointed at one part while the deck stayed
     * structurally valid and every check passed.
     *
     * It is the defect `nextNumber`'s own comment says it exists to prevent,
     * reached by a route the comment does not cover. A digit run too large to
     * count exactly is ignored rather than counted, which leaves the maximum
     * exact — and that is also what makes the free-number search terminate.
     */
    const pkg = await deck(ONE);
    // 2^53 exactly, which is the sharp case: `Number` reads it back precisely,
    // `max + 1` rounds straight back to it, and the name that produces is the
    // name already in the package. A longer run of nines is the same defect
    // with a different symptom — `max + 1` answers 1e+20, and the part named
    // after it is nonsense rather than a collision.
    for (const path of [
      "ppt/slides/slide9007199254740992.xml",
      "ppt/charts/chart9007199254740992.xml",
      "ppt/media/image9007199254740992.png",
    ]) {
      pkg.setBytes(path, new Uint8Array([1]));
    }
    expect(pkg.has(`ppt/slides/slide${pkg.nextSlideNumber()}.xml`), "the slide number is already taken").toBe(false);
    expect(pkg.has(`ppt/charts/chart${pkg.nextNumber("ppt/charts/chart")}.xml`)).toBe(false);
    expect(pkg.has(`ppt/media/image${pkg.nextMediaNumber()}.png`)).toBe(false);
    // And the ordinary contract is untouched: the highest real number plus one,
    // never filling a gap.
    expect(pkg.nextSlideNumber()).toBe(2);
  });

  it("refuses to name a part when the package's numbers leave nowhere safe to go", async () => {
    /**
     * The hole the rule above leaves. Ignoring a digit run too large to count
     * keeps the maximum exact — but a package holding the largest COUNTABLE
     * number and the one after it counts the first and ignores the second, so
     * "highest plus one" answers a name that is already there. That is the
     * collision the whole guard exists to prevent, one step further out.
     *
     * There is no larger safe number to hand back, so the answer is to refuse.
     * A deck reaching this has sixteen-digit part numbers and is not a deck.
     */
    const pkg = await deck(ONE);
    pkg.setBytes(`ppt/charts/chart${Number.MAX_SAFE_INTEGER}.xml`, new Uint8Array([1]));
    pkg.setBytes("ppt/charts/chart9007199254740992.xml", new Uint8Array([1]));
    expect(() => pkg.nextNumber("ppt/charts/chart")).toThrow(/too large to extend/);
  });

  it("reads a part's relationships once, not once per relationship added", async () => {
    /**
     * `addRel` re-read every `<Relationship>` in the part to find the highest
     * id, and the presentation's rels is the part that grows by one per merged
     * slide — so a merge was quadratic in the rows. Measured before the fix:
     * 250 rows took 353 ms and 2000 took 7364, eight times the work for
     * twenty-one times the time; after it, 4000 rows take 2705 ms.
     *
     * The assertion is on WORK rather than wall clock, which measures the
     * machine it happens to run on. One walk of the list is all this needs, and
     * the ids must still be distinct and ascending.
     */
    const pkg = await deck(ONE);
    const path = Pkg.relsPathFor("ppt/presentation.xml");
    const doc = await pkg.doc(path);
    let walks = 0;
    const real = doc.getElementsByTagNameNS.bind(doc);
    doc.getElementsByTagNameNS = ((ns: string, local: string) => {
      if (local === "Relationship") walks++;
      return real(ns, local);
    }) as typeof doc.getElementsByTagNameNS;

    const ids: string[] = [];
    for (let i = 0; i < 40; i++) ids.push(await pkg.addRel("ppt/presentation.xml", "http://example/t", `x${i}.xml`));
    expect(walks, "one walk per relationship added is what made a merge quadratic").toBeLessThanOrEqual(1);
    expect(new Set(ids).size, "two relationships were given one id").toBe(ids.length);
    const numbers = ids.map((id) => Number(id.slice(3)));
    expect([...numbers].sort((a, b) => a - b)).toEqual(numbers);
  });

  it("declares a content type once, without re-reading the whole list each time", async () => {
    // `addContentTypeOverride` asked `.some()` over a list it is itself
    // appending to, once per cloned chart or notes page.
    const pkg = await deck(ONE);
    const types = await pkg.doc("[Content_Types].xml");
    let walks = 0;
    const real = types.getElementsByTagNameNS.bind(types);
    types.getElementsByTagNameNS = ((ns: string, local: string) => {
      if (local === "Override") walks++;
      return real(ns, local);
    }) as typeof types.getElementsByTagNameNS;

    for (let i = 0; i < 40; i++) {
      await pkg.addContentTypeOverride(`/ppt/charts/chart${i}.xml`, "application/example");
    }
    expect(walks).toBeLessThanOrEqual(1);
    // Every one declared, and asking again does not add a second.
    await pkg.addContentTypeOverride("/ppt/charts/chart7.xml", "application/example");
    const declared = [...(await pkg.text("[Content_Types].xml")).matchAll(/PartName="([^"]+)"/g)].map((m) => m[1]);
    expect(declared.filter((n) => n === "/ppt/charts/chart7.xml")).toHaveLength(1);
    expect(declared).toContain("/ppt/charts/chart39.xml");
  });

  it("survives a round trip through base64", async () => {
    const pkg = await deck(ONE);
    const again = await Pkg.open(await pkg.toBase64());
    expect(await again.slidePaths()).toEqual(["ppt/slides/slide1.xml"]);
  });
});

describe("cloneSlide", () => {
  it("adds a part, a relationship, a content type and a slide id", async () => {
    const pkg = await deck(ONE);
    const copy = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });

    expect(copy).toBe("ppt/slides/slide2.xml");
    expect(await pkg.slidePaths()).toEqual(["ppt/slides/slide1.xml", "ppt/slides/slide2.xml"]);
    expect(await pkg.text("[Content_Types].xml")).toContain("/ppt/slides/slide2.xml");
    expect(pkg.has("ppt/slides/_rels/slide2.xml.rels")).toBe(true);
  });

  it("leaves a foreign custDataLst alone when it holds no tags of ours", async () => {
    /**
     * `dropInheritedTags` removes the template's `<p:tags>` from the copy, and
     * a `<p:custDataLst>` may hold customer data with no `<p:tags>` in it at
     * all — which `writeSlideTags`'s own fixture calls the ordinary shape of a
     * template built by another tool.
     *
     * The guard is a conjunction, and nothing exercised it: with `custData ||
     * tags` the clone reaches `removeChild(undefined)` on exactly this slide.
     * Found by `scripts/mutate-core.mjs`.
     */
    const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
    const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const pkg = await deck(ONE);
    pkg.setText(
      "ppt/slides/slide1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<p:sld ${P} ${A} ${R}><p:cSld><p:spTree>` +
        `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
        `</p:spTree>` +
        `<p:custDataLst><p:custData r:id="rId9"/></p:custDataLst>` +
        `</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
    );

    const copy = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });
    const xml = await pkg.text(copy);
    expect(xml, "the other tool's entry is not ours to remove").toContain('<p:custData r:id="rId9"');
    expect(xml.match(/<p:custDataLst/g) ?? []).toHaveLength(1);
  });

  it("gives every copy its own creation id", async () => {
    // office-js#6105: re-inserting a slide whose creation id is already in the
    // deck throws InvalidArgument on Windows desktop. Two copies sharing one is
    // the same collision one step earlier.
    const pkg = await deck(ONE);
    let n = 500;
    await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => ++n });
    await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => ++n });

    const ids = await Promise.all(["slide1", "slide2", "slide3"].map((s) => creationIdOf(pkg, `ppt/slides/${s}.xml`)));
    expect(new Set(ids).size).toBe(3);
  });

  it("refuses a creation id the deck is ALREADY using, and draws again", async () => {
    // The test above injects a counter and then asserts the counter is unique,
    // so the generator a real run uses was never in the assertion — and that
    // generator picked from 2^32 with nothing comparing the result against the
    // deck. A collision with the template's own slide, with the user's other
    // slides, or with an earlier copy in the same run is the state
    // office-js#6105 reports as InvalidArgument on Windows desktop.
    //
    // Handing over a value the deck already holds is the only way to reach that
    // branch deliberately: the odds of drawing one are about one in a hundred
    // thousand, which is exactly why it must not be left to the draw.
    const pkg = await deck(ONE); // slide1 carries creation id 111
    const offered = [111, 111, 777];
    const copy = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => offered.shift() ?? 0 });

    expect(await creationIdOf(pkg, copy)).toBe(777);
    expect(await creationIdOf(pkg, "ppt/slides/slide1.xml")).toBe(111);
  });

  it("honours a generator that keeps answering the same number", async () => {
    // Bounded rather than looped. A caller that answers 111 forever is being
    // deliberate, and spinning on it would be worse than taking the value —
    // the redraw exists for a random draw that got unlucky, not to overrule an
    // injected one.
    const pkg = await deck(ONE);
    const copy = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 111 });
    expect(await creationIdOf(pkg, copy)).toBe(111);
  });

  it("writes a creation id into a template that has none", async () => {
    const pkg = await deck([{ paragraphs: [["x"]] }]);
    expect(await creationIdOf(pkg, "ppt/slides/slide1.xml")).toBeUndefined();
    await setCreationId(pkg, "ppt/slides/slide1.xml", 4242);
    expect(await creationIdOf(pkg, "ppt/slides/slide1.xml")).toBe(4242);
  });

  it("gives every slide id in the deck a distinct value", async () => {
    const pkg = await deck([{ paragraphs: [["a"]] }, { paragraphs: [["b"]] }]);
    await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 9 });
    const pres = await pkg.doc("ppt/presentation.xml");
    const ids = elements(pres, P_NS, "sldId").map((e) => e.getAttribute("id"));
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("clones the notes page and repoints its back-reference", async () => {
    // A shared notes part means two slides editing one notes page, which looks
    // like the add-in overwriting the user's notes.
    const pkg = await deck([{ paragraphs: [["x"]], notes: true }]);
    await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 7 });

    expect(pkg.has("ppt/notesSlides/notesSlide2.xml")).toBe(true);
    expect(await pkg.text("ppt/slides/_rels/slide2.xml.rels")).toContain("notesSlide2.xml");
    expect(await pkg.text("ppt/notesSlides/_rels/notesSlide2.xml.rels")).toContain("slides/slide2.xml");
  });
});

describe("tags", () => {
  it("writes tags a reader can find again", async () => {
    const pkg = await deck(ONE);
    await writeSlideTags(pkg, "ppt/slides/slide1.xml", [
      [TAG_RUN, "run-1"],
      [TAG_BLOCK, "block-a"],
    ]);
    const back = await readSlideTags(pkg, "ppt/slides/slide1.xml");
    expect(back.get(TAG_RUN)).toBe("run-1");
    expect(back.get(TAG_BLOCK)).toBe("block-a");
  });

  it("survives the round trip a real merge takes", async () => {
    const pkg = await deck(ONE);
    await writeSlideTags(pkg, "ppt/slides/slide1.xml", [[TAG_RUN, "run-1"]]);
    const again = await Pkg.open(await pkg.toBase64());
    expect((await readSlideTags(again, "ppt/slides/slide1.xml")).get(TAG_RUN)).toBe("run-1");
  });

  it("never overwrites a tag part the template already had", async () => {
    // Writing tag1.xml blind destroys another tool's tags and every slide
    // pointing at them, silently.
    const pkg = await deck(ONE);
    pkg.setText("ppt/tags/tag1.xml", "<p:tagLst/>");
    expect(nextTagNumber(pkg)).toBe(2);
    await writeSlideTags(pkg, "ppt/slides/slide1.xml", [[TAG_RUN, "x"]]);
    expect(await pkg.text("ppt/tags/tag1.xml")).toBe("<p:tagLst/>");
  });

  it("appends to an existing tag list instead of adding a second one", async () => {
    // CT_CustomerDataList allows at most one <p:tags> child, so a slide that
    // already carries tags has to have ours merged in.
    const pkg = await deck(ONE);
    await writeSlideTags(pkg, "ppt/slides/slide1.xml", [["OTHER_TOOL", "keep me"]]);
    await writeSlideTags(pkg, "ppt/slides/slide1.xml", [[TAG_RUN, "run-2"]]);

    const back = await readSlideTags(pkg, "ppt/slides/slide1.xml");
    expect(back.get("OTHER_TOOL")).toBe("keep me");
    expect(back.get(TAG_RUN)).toBe("run-2");

    const doc = await pkg.doc("ppt/slides/slide1.xml");
    expect(elements(doc, P_NS, "tags")).toHaveLength(1);
  });

  // A real tag part declares the prefix; every one this engine sees comes from
  // tagPartXml or from PowerPoint, and both do.
  const part = (inner: string) => `<p:tagLst xmlns:p="${P_NS}">${inner}</p:tagLst>`;

  it("replaces our own key rather than duplicating it", () => {
    const merged = mergeTagPart(part(`<p:tag name="A" val="1"/><p:tag name="B" val="2"/>`), [["A", "9"]]);
    expect(merged).toContain('name="B" val="2"');
    expect(merged).toContain('name="A" val="9"');
    expect(merged.match(/name="A"/g)).toHaveLength(1);
  });

  it("does not re-escape a foreign value it keeps, however many merges it survives", () => {
    // The regex this replaced read attribute values as raw SOURCE, so an entity
    // was escaped again on every write: one merge turned `Ben & Jerry` into
    // `Ben &amp; Jerry` on screen, two into `Ben &amp;amp; Jerry`. The manual
    // promises other tools' tags are kept, and a value nobody can read back is
    // not kept.
    let xml = part(`<p:tag name="OTHER" val="Ben &amp; Jerry &lt;x&gt;"/>`);
    for (let i = 0; i < 3; i++) xml = mergeTagPart(xml, [[TAG_RUN, `run-${i}`]]);
    expect(xml).toContain('val="Ben &amp; Jerry &lt;x&gt;"');
  });

  it("keeps a foreign tag whatever legal spelling PowerPoint used", () => {
    // Attribute order, quote style and a separate closing tag are all legal and
    // all matched nothing before, so the foreign tag was silently DROPPED.
    for (const spelling of [
      `<p:tag val="keepme" name="OTHER"/>`,
      `<p:tag name="OTHER" val="keepme"></p:tag>`,
      `<p:tag name='OTHER' val='keepme'/>`,
      `<p:tag\n  name="OTHER"\n  val="keepme"/>`,
    ]) {
      const merged = mergeTagPart(part(spelling), [[TAG_RUN, "x"]]);
      expect(merged, spelling).toContain('name="OTHER" val="keepme"');
    }
  });

  it("escapes a value that would otherwise break the part", async () => {
    const pkg = await deck(ONE);
    await writeSlideTags(pkg, "ppt/slides/slide1.xml", [[TAG_RUN, `a"b&c<d`]]);
    expect((await readSlideTags(pkg, "ppt/slides/slide1.xml")).get(TAG_RUN)).toBe(`a"b&c<d`);
  });
});

/**
 * Four defects a bug hunt found in the package layer, each reproduced before it
 * was fixed. Three of them ship a file PowerPoint reads wrongly or not at all,
 * and none of them were visible from the engine's own output — the merge looked
 * like it had worked.
 */
describe("the package a merge actually writes", () => {
  it("declares xmlns:p14 exactly once on a creation id", async () => {
    // `createElementNS` binds the prefix and the serializer emits the
    // declaration, so setting it by hand as well produced a DUPLICATE
    // attribute. XML forbids that outright, so PowerPoint rejects the whole
    // package — and says nothing about which part.
    const pkg = await deck([{ paragraphs: [["x"]] }]);
    await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 4242 });
    const xml = await pkg.text("ppt/slides/slide2.xml");
    expect(xml.match(/xmlns:p14=/g) ?? []).toHaveLength(1);
    // And it must still parse, which is the thing the count is a proxy for.
    expect(elements(parseXml(xml), P14, "creationId")).toHaveLength(1);
  });

  it("puts the creation id on cSld, never inside the shape tree", async () => {
    // A slide whose spTree ends in its own extLst had the id appended THERE,
    // where PowerPoint does not look — so it invented one on open and two
    // copies were indistinguishable again, which is the collision the creation
    // id exists to prevent.
    const pkg = await deck([{ paragraphs: [["x"]] }]);
    const doc = await pkg.doc("ppt/slides/slide1.xml");
    const spTree = element(doc, P_NS, "spTree");
    spTree?.appendChild(doc.createElementNS(P_NS, "p:extLst"));
    await setCreationId(pkg, "ppt/slides/slide1.xml", 4242);

    const cSld = element(await pkg.doc("ppt/slides/slide1.xml"), P_NS, "cSld");
    const onCSld = children(cSld!, P_NS, "extLst").flatMap((e) => elements(e, P14, "creationId"));
    expect(onCSld).toHaveLength(1);
    expect(elements(element(await pkg.doc("ppt/slides/slide1.xml"), P_NS, "spTree")!, P14, "creationId")).toHaveLength(
      0,
    );
  });

  it("does not mistake a SHAPE's tags for the slide's", async () => {
    // The read this whole design rests on is slide.tags. A template holding a
    // shape with its own tag part — a sibling add-in's config, say — made
    // writeSlideTags append our metadata to that shape instead, so the slide
    // had no slide-level tags at all and undo could never find it.
    //
    // The shape's relationship has to RESOLVE for this to reproduce. A first
    // version of this test invented an rId that pointed at nothing, so the old
    // code failed to find a target, fell through to creating a fresh part, and
    // the guard passed against the unfixed file.
    const pkg = await deck([{ paragraphs: [["x"]] }]);
    pkg.setText("ppt/tags/tag1.xml", tagPartXml([["POWERCHART_CONFIG", "{}"]]));
    await pkg.addContentTypeOverride(
      "/ppt/tags/tag1.xml",
      "application/vnd.openxmlformats-officedocument.presentationml.tags+xml",
    );
    const shapeRel = await pkg.addRel(
      "ppt/slides/slide1.xml",
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags",
      "../tags/tag1.xml",
    );
    const doc = await pkg.doc("ppt/slides/slide1.xml");
    const sp = element(doc, P_NS, "sp");
    const nvPr = sp ? element(sp, P_NS, "nvPr") : undefined;
    const shapeList = doc.createElementNS(P_NS, "p:custDataLst");
    const shapeTags = doc.createElementNS(P_NS, "p:tags");
    shapeTags.setAttributeNS(R_NS, "r:id", shapeRel);
    shapeList.appendChild(shapeTags);
    nvPr?.appendChild(shapeList);

    await writeSlideTags(pkg, "ppt/slides/slide1.xml", [[TAG_RUN, "run-1"]]);

    // The slide answers for its own tag.
    expect((await readSlideTags(pkg, "ppt/slides/slide1.xml")).get(TAG_RUN)).toBe("run-1");
    // And the shape's part was not touched: it is not ours.
    expect(await pkg.text("ppt/tags/tag1.xml")).not.toContain(TAG_RUN);
  });

  it("gives a clone its own tag part instead of the template's", async () => {
    // The .rels are copied verbatim, so a template that already carried tags
    // handed every clone a relationship pointing at the TEMPLATE's part. All
    // the copies then wrote into one part: every merged slide read back the
    // LAST record's tags, and the user's own template was stamped as merge
    // output and matched by undo.
    const pkg = await deck([{ paragraphs: [["x"]] }]);
    await writeSlideTags(pkg, "ppt/slides/slide1.xml", [["OTHER_TOOL", "keep me"]]);

    const clones: string[] = [];
    for (const record of ["0", "1", "2"]) {
      const t = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 900 + clones.length });
      await writeSlideTags(pkg, t, [
        [TAG_RUN, "run-1"],
        [TAG_RECORD, record],
      ]);
      clones.push(t);
    }

    const records = await Promise.all(clones.map(async (t) => (await readSlideTags(pkg, t)).get(TAG_RECORD)));
    expect(records).toEqual(["0", "1", "2"]);

    const template = await readSlideTags(pkg, "ppt/slides/slide1.xml");
    expect(template.get("OTHER_TOOL")).toBe("keep me");
    expect(template.get(TAG_RUN)).toBeUndefined();
    expect(template.get(TAG_RECORD)).toBeUndefined();
  });
});

describe("where a part's relationships live", () => {
  it("puts them beside the part", () => {
    expect(Pkg.relsPathFor("ppt/slides/slide1.xml")).toBe("ppt/slides/_rels/slide1.xml.rels");
    expect(Pkg.relsPathFor("ppt/presentation.xml")).toBe("ppt/_rels/presentation.xml.rels");
  });

  it("handles a part at the package root", () => {
    // `lastIndexOf("/")` answers -1, and `slice(0, -1)` then drops the part's
    // last CHARACTER: the old answer was
    // `[Content_Types].xm/_rels/[Content_Types].xml.rels`, a plausible-looking
    // path to nowhere. Nothing calls it that way today — this closes it before
    // something does, because the failure is silent.
    expect(Pkg.relsPathFor("[Content_Types].xml")).toBe("_rels/[Content_Types].xml.rels");
  });
});

describe("a slide another tool has already written to", () => {
  const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

  /**
   * A slide carrying a `<p:custDataLst>` that holds only `<p:custData>` — no
   * `<p:tags>` of its own. `writeSlideTags`'s comment says a template built by
   * another tool routinely looks like this, and the branch that handles it was
   * the one branch in this file nothing exercised.
   *
   * It matters because `CT_CommonSlideData` allows at most one
   * `<p:custDataLst>`. Adding a second is not a tidiness problem: PowerPoint
   * opens the file as repaired and drops what it does not like, in the user's
   * own deck.
   */
  async function slideWithForeignCustData(): Promise<Pkg> {
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["hello"]] }]));
    pkg.setText(
      "ppt/slides/slide1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<p:sld ${P} ${A} ${R}><p:cSld><p:spTree>` +
        `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
        `</p:spTree>` +
        `<p:custDataLst><p:custData r:id="rId9"/></p:custDataLst>` +
        `</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
    );
    return pkg;
  }

  it("adds its tags to the list that is there, not a second one", async () => {
    const pkg = await slideWithForeignCustData();
    await writeSlideTags(pkg, "ppt/slides/slide1.xml", [[TAG_RUN, "run-1"]]);
    const xml = await pkg.text("ppt/slides/slide1.xml");
    expect(xml.match(/<p:custDataLst/g) ?? [], "a second custDataLst makes the file open as repaired").toHaveLength(1);
    // And the other tool's entry is still in it.
    expect(xml).toContain('<p:custData r:id="rId9"');
    expect((await readSlideTags(pkg, "ppt/slides/slide1.xml")).get(TAG_RUN)).toBe("run-1");
  });

  it("still round-trips through a real package", async () => {
    // Written, zipped, reopened — because the branch edits a live document and
    // the question is what the FILE ends up holding.
    const pkg = await slideWithForeignCustData();
    await writeSlideTags(pkg, "ppt/slides/slide1.xml", [[TAG_RUN, "run-2"]]);
    const again = await Pkg.open(await pkg.toBytes());
    expect((await readSlideTags(again, "ppt/slides/slide1.xml")).get(TAG_RUN)).toBe("run-2");
  });
});

describe("reading a part the run has edited but not written back", () => {
  it("answers with the edit, not the bytes on disk", async () => {
    /**
     * `Pkg` hands out parsed documents and only serialises them at the end, so
     * a reader that goes straight to the zip sees the version the file was
     * opened with. The changelog records this having bitten once already —
     * "three tests were passing on the version from disk" — and `maybeText` is
     * the reader my chart scan added a caller to.
     */
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["before"]] }]));
    const doc = await pkg.doc("ppt/slides/slide1.xml");
    const t = doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "t")[0];
    expect(t, "the fixture changed shape").toBeTruthy();
    t!.textContent = "after";
    expect(await pkg.maybeText("ppt/slides/slide1.xml")).toContain("after");
    expect(await pkg.maybeText("ppt/slides/slide1.xml")).not.toContain("before");
  });

  it("answers nothing for a part that is not there", async () => {
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["x"]] }]));
    expect(await pkg.maybeText("ppt/charts/chart1.xml")).toBeUndefined();
  });
});

describe("relating a part that had no relationships at all", () => {
  it("creates the rels file rather than failing", async () => {
    // `addRel` writes an empty `<Relationships>` when the part has none. Every
    // caller so far happened to work on a part that already had a rels file,
    // so the branch that creates one had never run.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["x"]] }]));
    const owner = "ppt/theme/theme1.xml";
    expect(pkg.has(Pkg.relsPathFor(owner)), "the fixture already gave theme1 a rels part").toBe(false);
    const rId = await pkg.addRel(owner, "http://example.invalid/rel", "../media/image1.png");
    expect(rId).toBe("rId1");
    expect(await pkg.relTarget(owner, rId)).toBe("ppt/media/image1.png");
  });
});

describe("a deck whose notes parts are numbered ahead of its slides", () => {
  /**
   * Part names in a package are arbitrary, and the slide and notes sequences
   * drift apart the moment a slide is deleted — so a one-slide deck can
   * perfectly well keep its notes in `notesSlide2.xml`. `cloneNotesSlide` named
   * the copy after the SLIDE number, which lands straight on that part.
   *
   * Nothing complains. `copyPart` overwrites silently and
   * `addContentTypeOverride` no-ops on an override that is already there, so the
   * package stays structurally valid and is wrong in two ways at once. Both were
   * reproduced on real bytes before the fix, and both are asserted here.
   */
  async function collidingDeck(): Promise<Pkg> {
    const bytes = await makeDeck([{ paragraphs: [["Hello {{First}}"]], notes: "Ring {{First}}" }]);
    const zip = await JSZip.loadAsync(bytes);
    const notes = await zip.file("ppt/notesSlides/notesSlide1.xml")!.async("uint8array");
    zip.remove("ppt/notesSlides/notesSlide1.xml");
    zip.file("ppt/notesSlides/notesSlide2.xml", notes);
    for (const path of ["ppt/slides/_rels/slide1.xml.rels", "[Content_Types].xml"]) {
      const s = await zip.file(path)!.async("string");
      zip.file(path, s.replaceAll("notesSlide1.xml", "notesSlide2.xml"));
    }
    return Pkg.open(await zip.generateAsync({ type: "uint8array" }));
  }

  it("gives the clone a notes page of its own rather than sharing the template's", async () => {
    const pkg = await collidingDeck();
    const template = await notesPathFor(pkg, "ppt/slides/slide1.xml");
    const clone = await cloneSlide(pkg, "ppt/slides/slide1.xml");
    expect(await notesPathFor(pkg, clone)).not.toBe(template);
  });

  it("does not put the FIRST record's notes on the second record's slide", async () => {
    /**
     * The consequence, and the one a user would actually meet. A clone sharing
     * the template's notes page has that page merged into it — so the next
     * clone copies notes whose placeholders are already gone, and ships the
     * record before it.
     */
    const pkg = await collidingDeck();
    const records = toRecordSet([["First"], ["Ada"], ["Grace"]]);
    const block = { id: "b", slides: [{ path: "ppt/slides/slide1.xml", seq: 1 }] };
    const result = await runPlan(pkg, buildPlan(block, records, { runId: "r" }), records);

    const second = result.slides[1];
    const notes = second === undefined ? undefined : await notesPathFor(pkg, second);
    expect(notes, "the second clone has no notes page").toBeDefined();
    expect((await pkg.doc(notes as string)).documentElement?.textContent).toBe("Ring Grace");
  });

  it("leaves no notes relationship pointing at a part that is not there", async () => {
    /**
     * The other consequence, and the one PowerPoint reports as a damaged file
     * without saying which part it could not find. The template is removed on
     * the way out — that is how the clones end up alone in the package — and
     * removing a slide takes its notes page with it. Shared, that page was the
     * clone's too.
     */
    const pkg = await collidingDeck();
    const records = toRecordSet([["First"], ["Ada"]]);
    const block = { id: "b", slides: [{ path: "ppt/slides/slide1.xml", seq: 1 }] };
    const result = await runPlan(pkg, buildPlan(block, records, { runId: "r" }), records);

    const keep = new Set(result.slides);
    for (const path of await pkg.slidePaths()) if (!keep.has(path)) await pkg.removeSlide(path);

    for (const slide of result.slides) {
      const notes = await notesPathFor(pkg, slide);
      expect(notes, `${slide} lost its notes page`).toBeDefined();
      expect(pkg.has(notes as string), `${slide} points at a notes part that is gone`).toBe(true);
    }
  });
});

describe("a template slide someone has commented on", () => {
  /**
   * A comment hangs off the SLIDE, so the wholesale rels copy a clone starts
   * from hands every copy a relationship to the TEMPLATE's comment part. Three
   * slides, one `modernComment_101_AEAB9DA1.xml` — measured before the fix.
   *
   * That puts a reviewer's "check this with Legal" on all 240 merged slides, as
   * one shared thread. Copying the part per clone would be worse rather than
   * better: the same note 240 times, deliberately. A comment is an annotation
   * about the template, not content the template produces.
   *
   * And dropping them is what makes the two template routes AGREE. On a 1.10
   * host `exportAsBase64Presentation` drops comments and `ppt/authors.xml`
   * outright (office-js#6867 — measured on this host 2026-08-28, four comment
   * parts in and none out), so the subset route already produced comment-free
   * clones while the file route produced shared ones. Two routes, two different
   * decks, from one template.
   */
  const MODERN = "http://schemas.microsoft.com/office/2018/10/relationships/comments";
  const CLASSIC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

  async function deckWithComment(relType: string, part: string): Promise<Pkg> {
    const bytes = await makeDeck([{ paragraphs: [["Hello {{First}}"]] }, { paragraphs: [["after"]] }]);
    const zip = await JSZip.loadAsync(bytes);
    zip.file(part, '<?xml version="1.0"?><cm/>');
    const rels = await zip.file("ppt/slides/_rels/slide1.xml.rels")!.async("string");
    zip.file(
      "ppt/slides/_rels/slide1.xml.rels",
      rels.replace(
        "</Relationships>",
        `<Relationship Id="rIdC" Type="${relType}" Target="../${part.replace("ppt/", "")}"/></Relationships>`,
      ),
    );
    return Pkg.open(await zip.generateAsync({ type: "uint8array" }));
  }

  const commentTargets = async (pkg: Pkg, slide: string): Promise<string[]> => {
    const path = Pkg.relsPathFor(slide);
    if (!pkg.has(path)) return [];
    const doc = await pkg.doc(path);
    return Array.from(doc.getElementsByTagName("Relationship"))
      .filter((r) => r.getAttribute("Type") === MODERN || r.getAttribute("Type") === CLASSIC)
      .map((r) => r.getAttribute("Target") ?? "");
  };

  it("does not put the template's comment on every copy", async () => {
    const pkg = await deckWithComment(MODERN, "ppt/comments/modernComment_101_AEAB9DA1.xml");
    const a = await cloneSlide(pkg, "ppt/slides/slide1.xml");
    const b = await cloneSlide(pkg, "ppt/slides/slide1.xml");
    expect(await commentTargets(pkg, a), "the first copy inherited it").toEqual([]);
    expect(await commentTargets(pkg, b), "the second copy inherited it").toEqual([]);
    // And the template keeps its own. This drops a copy's inherited reference,
    // never the user's comment.
    expect(await commentTargets(pkg, "ppt/slides/slide1.xml")).toHaveLength(1);
  });

  it("does the same for a CLASSIC comment part, not just the web's modern one", async () => {
    // PowerPoint on the web writes `modernComment_<id>_<hash>.xml` under a
    // Microsoft namespace; desktop has written `commentN.xml` under the
    // OpenXML one for years. A rule that knows only the spelling in front of it
    // is a rule that works on one host.
    const pkg = await deckWithComment(CLASSIC, "ppt/comments/comment1.xml");
    const clone = await cloneSlide(pkg, "ppt/slides/slide1.xml");
    expect(await commentTargets(pkg, clone)).toEqual([]);
  });

  it("takes the comment part away with the slide it belonged to", async () => {
    /**
     * The other half. A comment belongs to ONE slide and is unreachable once
     * that slide is gone — and now that a clone no longer references the
     * template's part, removing the template on the way out would strand it: a
     * part with a content-type override and nothing pointing at it.
     */
    const part = "ppt/comments/modernComment_101_AEAB9DA1.xml";
    const pkg = await deckWithComment(MODERN, part);
    expect(pkg.has(part)).toBe(true);
    await pkg.removeSlide("ppt/slides/slide1.xml");
    expect(pkg.has(part), "the comment part outlived its slide").toBe(false);
  });
});

describe("a tag value's whitespace survives a merge", () => {
  /**
   * An XML parser NORMALISES an attribute value: a literal newline, carriage
   * return or tab inside one is read back as a SPACE. `xmlAttr` escaped the
   * five markup characters and wrote those three literally, so they were gone
   * after one merge — and stable at the wrong value afterwards, which is the
   * shape nobody reports.
   *
   * It cannot reach our own tags; a run id and a record number have no
   * whitespace. It reaches a FOREIGN tag, which `docs/MANUAL.md` promises
   * survives, and which is the whole reason `mergeTagPart` keeps what it does
   * not own.
   */
  const read = (xml: string, name: string): string | undefined =>
    elements(parseXml(xml), P_NS, "tag")
      .find((t) => t.getAttribute("name") === name)
      ?.getAttribute("val") ?? undefined;

  it.each([
    ["a newline", "line1\nline2"],
    ["a carriage return", "a\rb"],
    ["a tab", "a\tb"],
    ["markup, still", 'Ben & Jerry <"x">'],
  ])("keeps %s through two merges", (_label, value) => {
    // Two, because the first write is what loses it and the second is what
    // would have made a re-escaping bug visible instead.
    let xml = tagPartXml([["OTHER", value]]);
    xml = mergeTagPart(xml, [[TAG_RUN, "run-1"]]);
    xml = mergeTagPart(xml, [[TAG_RUN, "run-2"]]);
    expect(read(xml, "OTHER")).toBe(value);
    expect(read(xml, TAG_RUN), "our own tag stopped round-tripping").toBe("run-2");
  });
});

describe("removing a slide leaves nothing of it behind", () => {
  /**
   * A part with a content-type override and nothing pointing at it is the shape
   * this file already chases for notes pages and for comments. Tag parts were
   * not on the list: `writeSlideTags` writes `ppt/tags/tagN.xml` per slide, one
   * slide points at it, and `orphanedParts` collected only charts and diagrams
   * — so every removed slide left its tags behind.
   *
   * It reaches further than a swept preview. On the `file` route the package IS
   * the user's whole presentation and every slide that is not a clone is
   * removed from it, so a deck whose slides carry tags — this add-in's own from
   * a previous merge, or another add-in's — shipped one orphan per slide back
   * into their deck.
   *
   * Swept rather than listed, because the part type that was missed is by
   * definition the one nobody would write an assertion for.
   */
  const OWNED = ["ppt/charts/", "ppt/diagrams/", "ppt/embeddings/", "ppt/notesSlides/", "ppt/tags/"];

  async function reach(pkg: Pkg, from: string): Promise<Set<string>> {
    const seen = new Set<string>();
    const queue = [from];
    while (queue.length) {
      for (const next of await pkg.relatedParts(queue.shift() as string)) {
        if (seen.has(next)) continue;
        seen.add(next);
        queue.push(next);
      }
    }
    return seen;
  }

  it("no part of a removed slide survives it, of any kind", async () => {
    const pkg = await Pkg.open(
      await makeDeck([
        {
          paragraphs: [["{{Name}}"]],
          notes: "Call {{Name}} afterwards",
          chart: { title: "{{Name}}", categories: ["{{Name}}", "b"], workbook: ["{{Name}}"], values: ["1", "42"] },
          smartArt: ["{{Name}}", "second"],
        },
        { paragraphs: [["after"]] },
      ]),
    );
    const records = toRecordSet([["Name"], ["Ada"], ["Bo"], ["Cy"]]);
    const block = { id: "r", slides: [{ path: "ppt/slides/slide1.xml", seq: 1, fields: [] }] };
    const out = await runPlan(pkg, buildPlan(block, records, { runId: "r" }), records);
    const before = pkg.partNames().length;

    for (const slide of out.slides) await pkg.removeSlide(slide);

    // Alive: anything the presentation or a surviving slide can still reach,
    // plus the `.rels` companion of anything alive — a companion belongs to its
    // owner and is not referred to by anybody.
    const alive = new Set<string>(["ppt/presentation.xml", "[Content_Types].xml", "_rels/.rels"]);
    for (const p of await reach(pkg, "ppt/presentation.xml")) alive.add(p);
    for (const slide of await pkg.slidePaths()) {
      alive.add(slide);
      for (const p of await reach(pkg, slide)) alive.add(p);
    }
    for (const p of [...alive]) alive.add(Pkg.relsPathFor(p));

    const stranded = pkg
      .partNames()
      .filter((p) => !alive.has(p) && OWNED.some((prefix) => p.startsWith(prefix)))
      .sort();

    expect(stranded, "left behind with a content-type override and nothing pointing at it").toEqual([]);
    expect(before - pkg.partNames().length, "the removal freed nothing, so the sweep proves nothing").toBeGreaterThan(
      20,
    );
  });
});

describe("a creation id the template carries in the wrong place", () => {
  /**
   * `setCreationId`'s append path is scoped to `cSld` and says why: a slide
   * whose shape tree ends in its own `<p:extLst>` had the id appended THERE,
   * where PowerPoint does not look for a slide's id, so it invented one on open
   * and two copies were indistinguishable — the collision that file exists to
   * avoid.
   *
   * The UPDATE path was not scoped. It took the first `p14:creationId`
   * anywhere in the part, so a stray one inside the shape tree was updated
   * instead, the function returned, and the slide was left with no id of its
   * own — the same failure reached from the other side, with `creationIdOf`
   * reporting that the stamp had worked.
   *
   * Not hypothetical. The comment in that file records that an older version of
   * it put ids exactly there, so a deck merged by that version carries one, and
   * using a merged deck as a template is an ordinary thing to do.
   */
  const CREATION_ID_URI = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}";

  /** A slide carrying a creation id inside its shape tree, and none of its own. */
  async function deckWithStrayId(value: string): Promise<Pkg> {
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["a"]] }, { paragraphs: [["b"]] }]));
    const doc = await pkg.doc("ppt/slides/slide1.xml");
    const cSld = element(doc, P_NS, "cSld") as Element;
    const spTree = child(cSld, P_NS, "spTree") as Element;
    const extLst = doc.createElementNS(P_NS, "p:extLst");
    const ext = doc.createElementNS(P_NS, "p:ext");
    ext.setAttribute("uri", CREATION_ID_URI);
    const stray = doc.createElementNS(P14, "p14:creationId");
    stray.setAttribute("val", value);
    ext.appendChild(stray);
    extLst.appendChild(ext);
    spTree.appendChild(extLst);
    return pkg;
  }

  const idsIn = (doc: Document): { val: string | null; under: string }[] =>
    Array.from(doc.getElementsByTagNameNS(P14, "creationId")).map((n) => ({
      val: n.getAttribute("val"),
      under: (n.parentNode?.parentNode?.parentNode as Element | null)?.localName ?? "?",
    }));

  it("stamps the SLIDE, not the stray one", async () => {
    const pkg = await deckWithStrayId("111111");
    await setCreationId(pkg, "ppt/slides/slide1.xml", 999999);

    const ids = idsIn(await pkg.doc("ppt/slides/slide1.xml"));
    expect(ids, "the stamp landed where PowerPoint does not look").toContainEqual({ val: "999999", under: "cSld" });
    // The stray is left where it is: deleting content out of somebody's deck to
    // tidy up is a bigger decision than this function is making, and an ignored
    // extension in a shape tree costs nothing.
    expect(ids).toContainEqual({ val: "111111", under: "spTree" });
  });

  it("and reads back the slide's own", async () => {
    // The diagnostic agreed with the broken stamp before, which is what made it
    // silent: `creationIdOf` returned the value it had just written into the
    // wrong element.
    const pkg = await deckWithStrayId("111111");
    await setCreationId(pkg, "ppt/slides/slide1.xml", 999999);
    expect(await creationIdOf(pkg, "ppt/slides/slide1.xml")).toBe(999999);
  });

  it("still updates in place when the id IS the slide's own", async () => {
    // The other half: scoping the search must not turn every stamp into an
    // append, or a slide accumulates ids and the first one wins.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["a"]] }, { paragraphs: [["b"]] }]));
    await setCreationId(pkg, "ppt/slides/slide1.xml", 111);
    await setCreationId(pkg, "ppt/slides/slide1.xml", 222);
    const ids = idsIn(await pkg.doc("ppt/slides/slide1.xml"));
    expect(ids).toEqual([{ val: "222", under: "cSld" }]);
  });
});

describe("a part that arrives with a byte order mark", () => {
  /**
   * OPC permits a BOM on an XML part and .NET's default `UTF8Encoding` emits
   * one, so any deck from a third-party generator built on it carries one on
   * every part that generator wrote. PowerPoint opens such a deck without a
   * murmur.
   *
   * `@xmldom/xmldom` does not: a leading `U+FEFF` puts the XML declaration at
   * position 1 and it throws `processing instruction at position 1 is an xml
   * declaration which is only at the start of the document`. JSZip's
   * `async("string")` hands the character through — it decodes UTF-8 and has
   * no opinion about what the first code point means — so the mark reaches the
   * parser as content and the merge dies on the first slide it reads, naming
   * neither the part nor the reason a user could act on.
   */
  async function deckWithBom(part: string): Promise<Uint8Array> {
    const zip = await JSZip.loadAsync(await makeDeck([{ paragraphs: [["Hello {{Name}}"]] }]));
    const file = zip.file(part);
    if (!file) throw new Error(`the fixture has no ${part}, so this test proves nothing`);
    zip.file(part, `\uFEFF${await file.async("string")}`);
    return zip.generateAsync({ type: "uint8array" });
  }

  it("reads a slide whose markup starts with one", async () => {
    const pkg = await Pkg.open(await deckWithBom("ppt/slides/slide1.xml"));
    // The mark really is still in the bytes — if JSZip ever started stripping
    // it, this test would pass while proving nothing.
    expect((await pkg.text("ppt/slides/slide1.xml")).charCodeAt(0)).toBe(0xfeff);
    expect(elements(await pkg.doc("ppt/slides/slide1.xml"), P_NS, "cSld")).toHaveLength(1);
  });

  it("merges a deck whose slide, presentation and content types all carry one", async () => {
    // Every part a merge has to parse, marked at once, which is what a whole
    // deck written by such a generator looks like.
    let bytes = await makeDeck([{ paragraphs: [["Hello {{Name}}"]] }]);
    for (const part of ["ppt/slides/slide1.xml", "ppt/presentation.xml", "[Content_Types].xml"]) {
      const zip = await JSZip.loadAsync(bytes);
      const file = zip.file(part);
      if (!file) throw new Error(`the fixture has no ${part}, so this test proves nothing`);
      zip.file(part, `\uFEFF${await file.async("string")}`);
      bytes = await zip.generateAsync({ type: "uint8array" });
    }
    const pkg = await Pkg.open(bytes);
    const prepared = await prepareBlock(pkg, { from: 1, to: 1, offsetInPackage: 0 }, "bom");
    if (!prepared.ok) throw new Error(`refused: ${prepared.why}`);
    const records = toRecordSet([["Name"], ["Ada"], ["Bo"]]);
    const result = await runPlan(pkg, buildPlan(prepared.block, records, { runId: "bom" }), records, {});
    expect(result.slides).toHaveLength(2);
    const said = await Promise.all(
      result.slides.map(async (s) => elements(await pkg.doc(s), A_NS, "t")[0]?.textContent ?? ""),
    );
    expect(said.join("|")).toContain("Ada");
  });
});

describe("a slide whose <p:tags> leads nowhere", () => {
  const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

  /**
   * A slide-level `<p:tags>` reference that does not resolve to a part, in the
   * two shapes a deck can carry it.
   *
   * `"no-part"` — the relationship is there and names a part the package does
   * not hold. PowerPoint writes this itself when it repairs a file: the part
   * goes and the reference stays.
   *
   * `"no-rel"` — the markup names a relationship id the slide's `.rels` does
   * not have, which is what a tool that deleted a relationship without
   * touching the markup leaves behind.
   *
   * `readSlideTags` degrades on both and answers an empty map. `writeSlideTags`
   * did not: it THREW on the first, because `Pkg.text` throws by name for a
   * missing part, and it silently appended a SECOND `<p:tags>` on the second.
   * `CT_CustomerDataList` caps that list at one, and `readSlideTags` reads the
   * first child — so the run's own tag became invisible to every reader of it,
   * which is the read undo depends on to find the slides it made.
   */
  async function slideWithBrokenTagRef(shape: "no-part" | "no-rel"): Promise<Pkg> {
    const zip = await JSZip.loadAsync(await makeDeck([{ paragraphs: [["hello"]] }]));
    zip.file(
      "ppt/slides/slide1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<p:sld ${P} ${A} ${R}><p:cSld><p:spTree>` +
        `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
        `</p:spTree>` +
        `<p:custDataLst><p:custData r:id="rId9"/><p:tags r:id="rId50"/></p:custDataLst>` +
        `</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
    );
    if (shape === "no-part") {
      const rels = zip.file("ppt/slides/_rels/slide1.xml.rels");
      if (!rels) throw new Error("the fixture wrote no slide rels, so this test proves nothing");
      zip.file(
        "ppt/slides/_rels/slide1.xml.rels",
        (await rels.async("string")).replace(
          "</Relationships>",
          `<Relationship Id="rId50" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"` +
            ` Target="../tags/tagGone.xml"/></Relationships>`,
        ),
      );
    }
    return Pkg.open(await zip.generateAsync({ type: "uint8array" }));
  }

  for (const shape of ["no-part", "no-rel"] as const) {
    it(`writes tags a reader can find again (${shape})`, async () => {
      const pkg = await slideWithBrokenTagRef(shape);
      // The reader's answer first, so the two halves of the pair are on record
      // together: it degrades, and the writer must not do worse than that.
      expect(await readSlideTags(pkg, "ppt/slides/slide1.xml")).toEqual(new Map());
      await writeSlideTags(pkg, "ppt/slides/slide1.xml", [[TAG_RUN, "run-1"]]);

      const again = await Pkg.open(await pkg.toBytes());
      const cSld = element(await again.doc("ppt/slides/slide1.xml"), P_NS, "cSld");
      const custDataLst = child(cSld as Element, P_NS, "custDataLst");
      expect(
        children(custDataLst as Element, P_NS, "tags"),
        "CT_CustomerDataList allows one <p:tags>; a second makes the run's own tag unreadable",
      ).toHaveLength(1);
      // The other tool's entry in the same list is not collateral.
      expect(children(custDataLst as Element, P_NS, "custData")).toHaveLength(1);
      expect((await readSlideTags(again, "ppt/slides/slide1.xml")).get(TAG_RUN)).toBe("run-1");
    });
  }
});

describe("a tag value carrying a character XML cannot hold", () => {
  /**
   * `mergeTagPart` carries a FOREIGN add-in's tags through a merge, which is a
   * promise `docs/MANUAL.md` makes — and the value comes from a template deck
   * this engine did not write. `xmlAttr` escapes the five markup characters and
   * the whitespace an attribute parser normalises, and stops there.
   *
   * The characters XML 1.0 forbids outright have no escape at all: `&#11;` is
   * exactly as ill-formed as the byte. So a value holding one produced a tag
   * part PowerPoint refuses, on every merged slide, reported as a damaged file
   * with nothing naming the cause.
   *
   * The rule already existed for slide TEXT and lived in `merge/text.ts`, one
   * layer above the only other place this codebase builds XML by
   * concatenation. It is in `pptx/xml.ts` now, where both readers can share it
   * — a second copy is how the two would come to disagree about what XML can
   * carry.
   */
  const FORBIDDEN = (xml: string) =>
    [...xml].filter((c) => {
      const n = c.charCodeAt(0);
      return (n < 0x20 && n !== 9 && n !== 10 && n !== 13) || n === 0xfffe || n === 0xffff;
    });

  it("writes a part a conforming parser will take", () => {
    for (const [what, value] of [
      ["a NUL", "a\u0000b"],
      ["a vertical tab", "a\u000bb"],
      ["a form feed", "a\u000cb"],
      ["U+FFFE", "p\ufffeq"],
    ] as const) {
      const xml = tagPartXml([["FOREIGN", value]]);
      expect(FORBIDDEN(xml), what).toEqual([]);
      expect(() => parseXml(xml), what).not.toThrow();
    }
  });

  it("keeps a lone surrogate out, and an astral character in", () => {
    // An unpaired half is already broken text and cannot be written at all; a
    // well-formed pair is one ordinary code point above FFFF and must survive,
    // which is the distinction `\p{Surrogate}` under the `u` flag draws.
    expect(tagPartXml([["FOREIGN", "x\ud800y"]])).not.toContain("\ud800");
    expect(tagPartXml([["FOREIGN", `ok ${String.fromCodePoint(0x1f600)}`]])).toContain(String.fromCodePoint(0x1f600));
  });

  it("still escapes the markup and the whitespace it always did", () => {
    // The neighbouring rule this must not disturb. A tab, a newline and a
    // carriage return are LEGAL XML and are escaped for a different reason: an
    // attribute parser normalises them to spaces, so writing them literally
    // loses them.
    const xml = tagPartXml([["FOREIGN", `a<b>&"c'\td\ne`]]);
    expect(xml).toContain("&lt;");
    expect(xml).toContain("&amp;");
    expect(xml).toContain("&quot;");
    expect(xml).toContain("&#9;");
    expect(xml).toContain("&#10;");
    const round = parseXml(xml);
    expect(round.getElementsByTagName("p:tag")[0]?.getAttribute("val")).toBe(`a<b>&"c'\td\ne`);
  });
});

describe("sweeping the template block off a long deck", () => {
  /**
   * `removeSlide` used to resolve every `<p:sldId>` in the deck through
   * `relTarget`, which re-walks the presentation's whole relationship list for
   * one id — so a single removal cost `deck x deck`, and `src/office/merge.ts`
   * removes the template block one slide at a time. Measured on a 100-slide
   * user deck plus 400 clones: 4.5 SECONDS of blocking work in a task-pane
   * WebView, after the merge had already finished and with nothing on screen to
   * say why.
   *
   * Asserted as work rather than as wall clock, because a timing threshold on a
   * shared runner is a test that fails for reasons nobody can act on. The
   * property is that resolving the id list is not a per-slide question: the
   * relationships are read once for the whole sweep. Pre-fix this counts one
   * `relTarget` call per `<p:sldId>` per removal — 4 950 of them for the deck
   * below — and the number is what the seconds were made of.
   */
  it("does not re-resolve the whole slide id list once per slide", async () => {
    const pkg = await Pkg.open(await makeDeck(Array.from({ length: 100 }, () => ({ paragraphs: [["a"]] }))));
    const paths = await pkg.slidePaths();
    const seen = vi.spyOn(pkg, "relTarget");
    for (const path of paths.slice(0, 50)) await pkg.removeSlide(path);
    expect(await pkg.slidePaths()).toHaveLength(50);
    expect(seen.mock.calls.length, "the id list is being re-resolved per slide").toBeLessThanOrEqual(paths.length);
    seen.mockRestore();
  });
});
