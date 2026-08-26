import { describe, expect, it } from "vitest";
import { toRecordSet } from "../src/core/data/recordset.js";
import { buildPlan, type Block } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import { creationIdOf, notesPathFor } from "../src/core/pptx/clone.js";
import { Pkg } from "../src/core/pptx/pkg.js";
import { TAG_RECORD, TAG_RUN, readSlideTags } from "../src/core/pptx/tags.js";
import { A_NS, elements } from "../src/core/pptx/xml.js";
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
