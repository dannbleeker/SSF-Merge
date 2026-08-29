import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { cloneSlide, creationIdOf, notesPathFor, setCreationId } from "../src/core/pptx/clone.js";
import { buildPlan } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import { toRecordSet } from "../src/core/data/recordset.js";
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
