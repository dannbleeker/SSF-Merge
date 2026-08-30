import type JSZip from "jszip";
import { PKG_REL_NS, elements, parseXml } from "../../src/core/pptx/xml.js";

/**
 * Every relationship in the package that names a part which is not there.
 *
 * The one check that stands for a whole round: PowerPoint reports a package it
 * cannot resolve as damaged, repairs it silently, and drops whatever it chose
 * to drop. A merge that produces one of these looks like it worked right up
 * until somebody opens the file.
 *
 * Shared rather than copied. It was written for `test-kit.test.ts` and the
 * listing's demo deck needs exactly the same answer, and the off-by-one below
 * is the kind of thing that gets fixed in one copy and not the other.
 */
export async function danglingRels(zip: JSZip): Promise<string[]> {
  const names = new Set(Object.keys(zip.files));
  const dangling: string[] = [];
  for (const name of names) {
    if (!name.endsWith(".rels")) continue;
    // The package's own `_rels/.rels` has no directory in front of it, and a
    // bare `indexOf` answers -1 there — which silently produces a base of
    // "_rels/.rel" and reports every root relationship as dangling. `Pkg`
    // carries the same warning about the same off-by-one.
    const cut = name.indexOf("/_rels/");
    const base = cut < 0 ? "" : name.slice(0, cut);
    for (const rel of elements(parseXml((await zip.file(name)?.async("string")) ?? ""), PKG_REL_NS, "Relationship")) {
      if ((rel.getAttribute("TargetMode") ?? "") === "External") continue;
      const target = rel.getAttribute("Target") ?? "";
      if (/^[a-z]+:/.test(target)) continue;
      const segments = base.split("/").filter(Boolean);
      for (const seg of target.split("/")) {
        if (seg === "..") segments.pop();
        else if (seg !== ".") segments.push(seg);
      }
      if (!names.has(segments.join("/"))) dangling.push(`${name} -> ${target}`);
    }
  }
  return dangling;
}
