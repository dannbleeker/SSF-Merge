import { describe, expect, it } from "vitest";
import { cloneSlide, creationIdOf, setCreationId } from "../src/core/pptx/clone.js";
import { Pkg } from "../src/core/pptx/pkg.js";
import {
  TAG_BLOCK,
  TAG_RUN,
  mergeTagPart,
  nextTagNumber,
  readSlideTags,
  writeSlideTags,
} from "../src/core/pptx/tags.js";
import { P_NS, elements } from "../src/core/pptx/xml.js";
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

  it("replaces our own key rather than duplicating it", () => {
    const merged = mergeTagPart(`<p:tagLst><p:tag name="A" val="1"/><p:tag name="B" val="2"/></p:tagLst>`, [
      ["A", "9"],
    ]);
    expect(merged).toContain('name="B" val="2"');
    expect(merged).toContain('name="A" val="9"');
    expect(merged.match(/name="A"/g)).toHaveLength(1);
  });

  it("escapes a value that would otherwise break the part", async () => {
    const pkg = await deck(ONE);
    await writeSlideTags(pkg, "ppt/slides/slide1.xml", [[TAG_RUN, `a"b&c<d`]]);
    expect((await readSlideTags(pkg, "ppt/slides/slide1.xml")).get(TAG_RUN)).toBe(`a"b&c<d`);
  });
});
