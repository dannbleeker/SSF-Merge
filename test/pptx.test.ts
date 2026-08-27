import { describe, expect, it } from "vitest";
import { cloneSlide, creationIdOf, setCreationId } from "../src/core/pptx/clone.js";
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
import { P_NS, R_NS, children, element, elements, parseXml } from "../src/core/pptx/xml.js";

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
