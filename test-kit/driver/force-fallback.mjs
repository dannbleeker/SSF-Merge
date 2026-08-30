#!/usr/bin/env node
/**
 * Make every consumer take a modern chart's `mc:Fallback` branch.
 *
 * A modern chart is written as `<mc:AlternateContent>` holding an `<mc:Choice>`
 * — the real chartEx — and an `<mc:Fallback>` for hosts that cannot draw one.
 * SSF Merge replaces that fallback with a notice, because the picture
 * PowerPoint left there is the TEMPLATE's rendering, and on a mail merge that
 * is another recipient's figures under this recipient's name.
 *
 * The notice therefore renders only in PowerPoint 2013 and earlier, which is
 * why nobody had seen it until 2026-08-30. This is how it was seen without one.
 *
 * `mc:Choice` carries a `Requires` attribute naming namespace prefixes. A
 * consumer that cannot satisfy them skips the Choice and moves on. So repoint
 * the namespace at a URI nobody understands and every host — current PowerPoint
 * included — draws what an old one would.
 *
 * **Do not "force" it by deleting the Choice.** That is the obvious way and it
 * is wrong twice over. `mc:AlternateContent` must carry at least one Choice, so
 * the result is malformed: PowerPoint for the web opened such a file READ ONLY
 * with "We repaired your presentation. The original file is available." — which
 * looks exactly like the merged deck being damaged, the one finding this whole
 * test kit exists to catch. It is a fact about the mangling. The same run's
 * unmangled deck opened clean, which is the check to make before believing it.
 *
 * Usage: node test-kit/driver/force-fallback.mjs <in.pptx> [out.pptx]
 */
import { readFileSync, writeFileSync } from "node:fs";
import JSZip from "jszip";

const IN = process.argv[2];
const OUT = process.argv[3] ?? IN.replace(/\.pptx$/i, "") + "-fallback.pptx";
if (!IN) {
  console.error("usage: force-fallback.mjs <in.pptx> [out.pptx]");
  process.exit(2);
}

/** The namespace a modern chart's Choice requires. */
const CHARTEX = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";
/** One no consumer can satisfy, which is the whole trick. */
const NOBODY = "http://example.invalid/not-a-namespace-any-host-knows";

const zip = await JSZip.loadAsync(readFileSync(IN));
let moved = 0;
const touched = [];

for (const name of Object.keys(zip.files)) {
  if (!/^ppt\/slides\/slide\d+\.xml$/.test(name)) continue;
  const xml = await zip.file(name).async("string");
  // Only the declaration on the Choice itself. The chart part keeps its own
  // namespace, and anything else using it is left alone — this is meant to
  // change which BRANCH is taken, not what the branch contains.
  const out = xml.replace(/(<mc:Choice\b[^>]*xmlns:cx1=")([^"]*)(")/g, (whole, head, uri, tail) =>
    uri === CHARTEX ? `${head}${NOBODY}${tail}` : whole,
  );
  if (out !== xml) {
    moved++;
    touched.push(name.replace("ppt/slides/", ""));
    zip.file(name, out);
  }
}

if (moved === 0) {
  // Said loudly. A file that came back unchanged looks like a deck with no
  // modern chart and like an anchor that missed, and only one of those is worth
  // acting on.
  console.error(`no <mc:Choice> requiring chartex found in ${IN} — nothing was forced, and this file proves nothing`);
  process.exit(1);
}

writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer" }));
console.log(`forced ${moved} chart(s) onto the fallback: ${touched.join(", ")}`);
console.log(`WROTE: ${OUT}`);
console.log(`Open it in ANY PowerPoint. The merged copies should show the notice, not a chart and not a hole.`);
