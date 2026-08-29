import { describe, expect, it } from "vitest";
import { toRecordSet } from "../src/core/data/recordset.js";
import { buildPlan, type Block } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import { creationIdOf, notesPathFor } from "../src/core/pptx/clone.js";
import { Pkg } from "../src/core/pptx/pkg.js";
import { TAG_RECORD, TAG_RUN, readSlideTags } from "../src/core/pptx/tags.js";
import { A_NS, P_NS, R_NS, element, elements } from "../src/core/pptx/xml.js";
import { makeDeck } from "./fixtures/deck.js";

/** Every character the slide draws, in order. */
async function textOf(pkg: Pkg, path: string): Promise<string> {
  const doc = await pkg.doc(path);
  return elements(doc, A_NS, "t")
    .map((t) => t.textContent ?? "")
    .join("");
}

const records = toRecordSet([
  ["Name", "Revenue", "Notes"],
  ["Ada Lovelace", "1234567", "first"],
  ["Grace Hopper", "2140000", ""],
]);

/**
 * A two-slide block. The placeholders on slide 1 are SPLIT across runs, which
 * is what PowerPoint actually stores after an edit, and the whole reason the
 * replacement works on the paragraph rather than the node.
 */
async function template(): Promise<Pkg> {
  return Pkg.open(
    await makeDeck([
      { paragraphs: [["Hello {{Na", "me}}"], ["Revenue: {{Revenue|number}}"]], creationId: 11 },
      { paragraphs: [["Notes for {{Name}}: {{Notes}}"]], creationId: 22 },
    ]),
  );
}

const block: Block = {
  id: "block-a",
  slides: [
    { path: "ppt/slides/slide1.xml", seq: 1 },
    { path: "ppt/slides/slide2.xml", seq: 2, condition: "Notes" },
  ],
};

describe("runPlan", () => {
  it("produces a slide per planned step, appended after the template", async () => {
    const pkg = await template();
    const plan = buildPlan(block, records, { runId: "run-1" });
    let n = 0;
    const result = await runPlan(pkg, plan, records, { clone: { creationId: () => 900 + ++n } });

    // Ada gets both slides, Grace's Notes is empty so her second is skipped.
    expect(result.slides).toEqual(["ppt/slides/slide3.xml", "ppt/slides/slide4.xml", "ppt/slides/slide5.xml"]);
    expect(await pkg.slidePaths()).toHaveLength(5);
  });

  it("writes each record's values into its own copy", async () => {
    const pkg = await template();
    const plan = buildPlan(block, records, { runId: "run-1" });
    const result = await runPlan(pkg, plan, records);

    expect(await textOf(pkg, result.slides[0]!)).toBe("Hello Ada LovelaceRevenue: 1 234 567");
    expect(await textOf(pkg, result.slides[1]!)).toBe("Notes for Ada Lovelace: first");
    expect(await textOf(pkg, result.slides[2]!)).toBe("Hello Grace HopperRevenue: 2 140 000");
  });

  it("leaves the template slides exactly as they were", async () => {
    // Merging the template and cloning afterwards is the mistake this guards:
    // it would leave the template holding one record's values, which is the
    // template destroyed rather than used.
    const pkg = await template();
    const plan = buildPlan(block, records, { runId: "run-1" });
    await runPlan(pkg, plan, records);

    expect(await textOf(pkg, "ppt/slides/slide1.xml")).toBe("Hello {{Name}}Revenue: {{Revenue|number}}");
    expect(await textOf(pkg, "ppt/slides/slide2.xml")).toBe("Notes for {{Name}}: {{Notes}}");
  });

  it("tags every produced slide with the run and the record it came from", async () => {
    const pkg = await template();
    const plan = buildPlan(block, records, { runId: "run-1" });
    const result = await runPlan(pkg, plan, records);

    const tags = await readSlideTags(pkg, result.slides[2]!);
    expect(tags.get(TAG_RUN)).toBe("run-1");
    expect(tags.get(TAG_RECORD)).toBe("1");
  });

  it("leaves a SHAPE's own tag reference pointing where it pointed", async () => {
    // A copy must not inherit the template's BLOCK and SEQ tags — that is what
    // `dropInheritedTags` is for, and those live in the slide's own `<p:cSld>`.
    // An add-in's tags do not: they hang off a SHAPE, in
    // `<p:nvPr><p:custDataLst><p:tags r:id="…"/>`, and a deck touched by
    // think-cell carries exactly that on a hidden shape in every slide.
    //
    // Dropping every tag relationship took that one too, and the shape then
    // named a relationship that was gone. Which was the mild half: deleting a
    // relationship frees its ID, `writeSlideTags` takes the next free one for
    // this run's own tags, and the vendor's shape came out of the merge
    // pointing at SSF Merge's merge metadata. A reference that still resolves,
    // to somebody else's data — which is why this asserts the TARGET rather
    // than that the id is still there.
    // On BOTH template slides, so every copy is expected to carry it and the
    // assertion needs no per-slide bookkeeping to stay honest.
    const pkg = await Pkg.open(
      await makeDeck([
        { paragraphs: [["{{Name}}"]], shapeTags: true },
        { paragraphs: [["Second"]], shapeTags: true },
        { paragraphs: [["after"]] },
      ]),
    );
    const plan = buildPlan(block, records, { runId: "run-1" });
    const result = await runPlan(pkg, plan, records);

    for (const slide of result.slides) {
      const doc = await pkg.doc(slide);
      const refs = elements(doc, P_NS, "tags").map((t) => t.getAttributeNS(R_NS, "id") ?? t.getAttribute("r:id") ?? "");
      const targets = await Promise.all(refs.map((id) => pkg.relTarget(slide, id)));
      expect(targets, `${slide} names a tag relationship that is not there`).not.toContain(undefined);
      expect(targets, `${slide} lost the vendor's tag part`).toContain("ppt/tags/tag9.xml");
    }
  });

  it("gives every produced slide its own creation id", async () => {
    let n = 0;
    const pkg = await template();
    const plan = buildPlan(block, records, { runId: "run-1" });
    const result = await runPlan(pkg, plan, records, { clone: { creationId: () => 500 + ++n } });

    const ids = await Promise.all(result.slides.map((s) => creationIdOf(pkg, s)));
    expect(new Set([...ids, 11, 22]).size).toBe(5);
  });

  it("survives the round trip a real merge takes, bytes and back", async () => {
    // The deck is handed to PowerPoint as base64. Everything above is worth
    // nothing if it does not survive being written out and read again.
    const pkg = await template();
    const plan = buildPlan(block, records, { runId: "run-1" });
    const result = await runPlan(pkg, plan, records);

    const again = await Pkg.open(await pkg.toBase64());
    expect(await again.slidePaths()).toHaveLength(5);
    expect(await textOf(again, result.slides[0]!)).toBe("Hello Ada LovelaceRevenue: 1 234 567");
    expect((await readSlideTags(again, result.slides[0]!)).get(TAG_RUN)).toBe("run-1");
  });

  it("leaves the placeholder visible when the policy is keep and the cell is empty", async () => {
    const pkg = await template();
    const notesBlock: Block = { id: "b", slides: [{ path: "ppt/slides/slide2.xml", seq: 1 }] };
    const plan = buildPlan(notesBlock, records, { runId: "run-1", recordIndexes: [1] });
    const result = await runPlan(pkg, plan, records, { onEmpty: "keep" });

    expect(await textOf(pkg, result.slides[0]!)).toBe("Notes for Grace Hopper: {{Notes}}");
  });

  it("writes nothing for an empty cell when the policy is blank", async () => {
    const pkg = await template();
    const notesBlock: Block = { id: "b", slides: [{ path: "ppt/slides/slide2.xml", seq: 1 }] };
    const plan = buildPlan(notesBlock, records, { runId: "run-1", recordIndexes: [1] });
    const result = await runPlan(pkg, plan, records, { onEmpty: "blank" });

    expect(await textOf(pkg, result.slides[0]!)).toBe("Notes for Grace Hopper: ");
  });
});

describe("what a merged copy carries", () => {
  it("merges the notes page too, not just the slide", async () => {
    // A copy gets its own notes slide precisely so the copies can differ. Left
    // unmerged, a template whose speaker notes read "Call {{Name}} afterwards"
    // ships that text verbatim on every merged slide — in the presenter view
    // and on every printed handout. text.ts declares charts and SmartArt as the
    // deliberate exclusions; notes were never one of them.
    const pkg = await Pkg.open(
      await makeDeck([{ paragraphs: [["Hello {{Name}}"]], notes: "Call {{Name}} afterwards" }]),
    );
    const rows = toRecordSet([["Name"], ["Ada"], ["Grace"]]);
    const block: Block = { id: "b", slides: [{ path: "ppt/slides/slide1.xml", seq: 1 }] };
    const plan = buildPlan(block, rows, { runId: "run-1" });
    const result = await runPlan(pkg, plan, rows, { clone: { creationId: () => 900 } });

    const notes: string[] = [];
    for (const slide of result.slides) {
      const path = await notesPathFor(pkg, slide);
      notes.push(path ? await textOf(pkg, path) : "");
    }
    expect(notes.join(" | ")).toContain("Call Ada afterwards");
    expect(notes.join(" | ")).toContain("Call Grace afterwards");
    expect(notes.join(" | ")).not.toContain("{{Name}}");
  });
});

describe("what a long merge holds in memory", () => {
  it("does not keep one parsed document per output slide", async () => {
    // The cache is also the dirty-part set, so nothing ever left it: a run held
    // one live xmldom Document per output slide on top of the zip's own copy of
    // the same bytes. Measured on a 300-paragraph slide, 300 clones held 440 MB
    // of heap against 54 MB released and 400 clones held 591 MB against 54 —
    // flat rather than growing, which is the property that matters inside a
    // task-pane WebView.
    //
    // Counted rather than measured: a heap assertion would be flaky, and what
    // is actually being claimed is that the held count does not track the
    // record count.
    const held = async (n: number) => {
      const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["Hello {{Name}}"]], notes: "Notes {{Name}}" }]));
      const rows = toRecordSet([["Name"], ...Array.from({ length: n }, (_, i) => [`R${i}`])]);
      const block: Block = { id: "b", slides: [{ path: "ppt/slides/slide1.xml", seq: 1 }] };
      let id = 0;
      await runPlan(pkg, buildPlan(block, rows, { runId: "r" }), rows, { clone: { creationId: () => 900 + ++id } });
      return pkg.cachedParts();
    };
    expect(await held(20)).toBe(await held(2));
  });

  it("does not keep one parsed chart or diagram per output slide either", async () => {
    // The same property, on the parts a chart adds. Every copy gets its own
    // chart, its own workbook, its own SmartArt model and its own rendering —
    // four more parts per record — so a merge that held them would grow four
    // times faster than the one this file already measured. The original
    // measurement was taken on a slide with neither, so it could not have seen
    // this.
    const held = async (n: number) => {
      const pkg = await Pkg.open(
        await makeDeck([
          {
            paragraphs: [["Hello {{Name}}"]],
            chart: { title: "{{Name}}", categories: ["{{Name}}"], workbook: ["{{Name}}"] },
            smartArt: ["{{Name}}"],
          },
        ]),
      );
      const rows = toRecordSet([["Name"], ...Array.from({ length: n }, (_, i) => [`R${i}`])]);
      const block: Block = { id: "b", slides: [{ path: "ppt/slides/slide1.xml", seq: 1 }] };
      let id = 0;
      await runPlan(pkg, buildPlan(block, rows, { runId: "r" }), rows, { clone: { creationId: () => 900 + ++id } });
      return pkg.cachedParts();
    };
    expect(await held(20)).toBe(await held(2));
  });

  it("still writes every part it released", async () => {
    // Releasing serialises back into the zip first. If it did not, the whole
    // merge would come out as the untouched template with nothing to show for
    // it — which is the one outcome worse than being slow.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["Hello {{Name}}"]], notes: "Notes {{Name}}" }]));
    const rows = toRecordSet([["Name"], ["Ada"]]);
    const block: Block = { id: "b", slides: [{ path: "ppt/slides/slide1.xml", seq: 1 }] };
    const result = await runPlan(pkg, buildPlan(block, rows, { runId: "r" }), rows, {
      clone: { creationId: () => 900 },
    });
    const slide = result.slides[0] ?? "";
    // Read through text(), which goes to the zip for a part no longer cached.
    expect(await pkg.text(slide)).toContain("Hello Ada");
    const notes = await notesPathFor(pkg, slide);
    expect(await pkg.text(notes ?? "")).toContain("Notes Ada");
    // And the tags written after the merge survived the release too.
    expect((await readSlideTags(pkg, slide)).get(TAG_RUN)).toBe("r");
  });
});

describe("the package a merge hands to PowerPoint", () => {
  it("carries the copies and NOT the template it cloned from", async () => {
    // The run clones inside the package, so the package holds both. Inserted
    // whole, it would put the user's own placeholder slides back into their
    // deck after every merge. Removing them is ours to get right — the
    // alternative, naming only the copies through insertSlidesFromBase64's
    // sourceSlideIds, needs ids in the host's own spelling CONSTRUCTED for a
    // package not yet in the presentation, which no round has tested.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["Hello {{Name}}"]] }]));
    const rows = toRecordSet([["Name"], ["Ada"], ["Grace"]]);
    const block: Block = { id: "b", slides: [{ path: "ppt/slides/slide1.xml", seq: 1 }] };
    const result = await runPlan(pkg, buildPlan(block, rows, { runId: "r" }), rows, {
      clone: { creationId: () => 900 },
    });
    expect(result.slides).toHaveLength(2);

    await pkg.removeSlide("ppt/slides/slide1.xml");

    const left = await pkg.slidePaths();
    expect(left).toEqual(result.slides);
    expect(pkg.has("ppt/slides/slide1.xml")).toBe(false);

    // And it survives the round trip, which is what PowerPoint actually gets.
    const back = await Pkg.open(await pkg.toBase64());
    expect(await back.slidePaths()).toHaveLength(2);
    const texts = await Promise.all((await back.slidePaths()).map((s) => textOf(back, s)));
    expect(texts.sort()).toEqual(["Hello Ada", "Hello Grace"]);
  });

  it("leaves nothing behind that points at the slide it removed", async () => {
    // A dangling relationship or a content-type override for a part that is not
    // there is a package no reader will open, and it is the easy half to forget.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["a"]], notes: true }, { paragraphs: [["b"]] }]));
    await pkg.removeSlide("ppt/slides/slide1.xml");

    const back = await Pkg.open(await pkg.toBytes());
    expect(back.has("ppt/slides/slide1.xml")).toBe(false);
    expect(back.has("ppt/slides/_rels/slide1.xml.rels")).toBe(false);
    // Its notes page belonged to it and is unreachable now.
    expect(back.has("ppt/notesSlides/notesSlide1.xml")).toBe(false);

    const types = await back.text("[Content_Types].xml");
    expect(types).not.toContain("/ppt/slides/slide1.xml");
    expect(types).not.toContain("/ppt/notesSlides/notesSlide1.xml");

    const rels = await back.text(Pkg.relsPathFor("ppt/presentation.xml"));
    expect(rels).not.toContain("slides/slide1.xml");
    // The slide that stayed is untouched.
    expect(await back.slidePaths()).toEqual(["ppt/slides/slide2.xml"]);

    // And no <p:sldId> is left pointing at a relationship that is gone.
    //
    // slidePaths() cannot see this: it resolves each rId and SKIPS the ones
    // that answer nothing, so a dangling entry reads as a tidy deck while
    // presentation.xml references a relationship that does not exist — which
    // PowerPoint refuses. Proven by removing the id-list line from
    // `removeSlide`: without this assertion every other check still passed.
    const pres = await back.doc("ppt/presentation.xml");
    const list = element(pres, P_NS, "sldIdLst");
    const ids = list ? elements(list, P_NS, "sldId") : [];
    expect(ids).toHaveLength(1);
    for (const sldId of ids) {
      const rId = sldId.getAttributeNS(R_NS, "id") ?? sldId.getAttribute("r:id");
      expect(rId, "every sldId names a relationship").toBeTruthy();
      expect(await back.relTarget("ppt/presentation.xml", rId ?? ""), `${rId ?? ""} resolves`).toBeTruthy();
    }
  });
});
