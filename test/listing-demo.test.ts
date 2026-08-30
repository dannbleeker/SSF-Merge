import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Pkg } from "../src/core/pptx/pkg.js";
import { prepareBlock } from "../src/core/merge/prepare.js";
import { buildPlan } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import JSZip from "jszip";
import { parseDelimited, toRecordSet } from "../src/core/data/recordset.js";
import { danglingRels } from "./fixtures/dangling.js";

/**
 * The deck the Marketplace screenshots are taken from.
 *
 * `docs/listing/demo/` holds a template deck and the rows to merge into it,
 * built by `build-demo.py` and committed beside it. Nothing in the product
 * depends on them, so without this file they are two binaries that were correct
 * on the day somebody shot them.
 *
 * The failure this is really guarding against is specific and expensive: a
 * placeholder whose spelling does not match a column is NOT an error at merge
 * time. The engine leaves the field on the slide on purpose, because a slide
 * that looks finished and is not is worse than one that says what is missing.
 * That is the right behaviour in the product and a disaster in a store listing,
 * where it means shipping a screenshot with `{{Account}}` printed on the slide.
 * A person would find it by squinting at a picture. This finds it by asking.
 *
 * It runs the same four calls the task pane does rather than a wrapper written
 * for the test, which would only prove that the wrapper works.
 */

const DEMO = "docs/listing/demo";

/** Slides 2 and 3 are the template block; slide 1 is the deck's cover. */
const BLOCK = { from: 2, to: 3, offsetInPackage: 1 } as const;

async function mergeTheDemo() {
  const pkg = await Pkg.open(new Uint8Array(readFileSync(`${DEMO}/Quarterly-business-review.pptx`)));
  const prepared = await prepareBlock(pkg, BLOCK, "demo");
  if (!prepared.ok) throw new Error(`the demo template was refused: ${prepared.why}`);

  const records = toRecordSet(parseDelimited(readFileSync(`${DEMO}/rows.txt`, "utf8")));
  const plan = buildPlan(prepared.block, records, { runId: "demo" });
  const result = await runPlan(pkg, plan, records, {});
  return { pkg, prepared, records, plan, result };
}

describe("the listing's demo deck", () => {
  it("is a template block the engine accepts", async () => {
    const { prepared } = await mergeTheDemo();
    expect(prepared.ok).toBe(true);
  });

  it("has a column behind every placeholder on its slides", async () => {
    const { prepared, records } = await mergeTheDemo();
    const fields = prepared.ok ? prepared.fields : [];
    expect(fields.length, "the demo template has no placeholders at all").toBeGreaterThan(0);

    const columns = new Set(records.columns.map((c) => c.name));
    const orphans = fields.filter((f) => !columns.has(f));
    expect(orphans, "these would print as {{braces}} in a store screenshot").toEqual([]);
  });

  it("adds one set of slides per row", async () => {
    const { records, plan } = await mergeTheDemo();
    // Two template slides, three rows. The listing description says "one set of
    // slides per row", and a screenshot that shows otherwise contradicts it.
    expect(records.rows.length).toBe(3);
    expect(plan.steps.length).toBe(records.rows.length * 2);
  });

  it("merges into a package with no dangling relationships", async () => {
    // The one check that stands for the rest: PowerPoint reports a package it
    // cannot resolve as damaged and repairs it silently. A repair prompt in the
    // middle of a screenshot session is a wasted session.
    //
    // This asserted `result.slides.length > 0` when it was written, which is
    // not a relationship check at all — the same shape of mistake as a comment
    // promising more than the code does. It now runs the real one.
    const { pkg, result } = await mergeTheDemo();
    const keep = new Set(result.slides);
    for (const path of await pkg.slidePaths()) if (!keep.has(path)) await pkg.removeSlide(path);
    expect(await danglingRels(await JSZip.loadAsync(await pkg.toBytes()))).toEqual([]);
  });
});
