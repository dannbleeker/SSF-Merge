import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { buildPlan } from "../src/core/merge/plan.js";
import { runPlan } from "../src/core/merge/run.js";
import { toRecordSet } from "../src/core/data/recordset.js";
import { Pkg } from "../src/core/pptx/pkg.js";
import { A_NS, elements } from "../src/core/pptx/xml.js";
import { makeDeck, xfrm } from "./fixtures/deck.js";
import { makeResolver } from "../src/core/merge/resolve.js";
import { imageMode } from "../src/core/merge/images.js";

const WIDE = new Uint8Array(readFileSync("test/fixtures/wide.png")); // 64 x 32
const TALL = new Uint8Array(readFileSync("test/fixtures/tall.jpg")); // 30 x 90

/** A one-slide template whose only text is an image placeholder. */
async function template(field: string, box?: string): Promise<Pkg> {
  return Pkg.open(await makeDeck([{ paragraphs: [[field]], ...(box ? { box } : {}) }, { paragraphs: [["after"]] }]));
}

const BLOCK = { id: "b", slides: [{ path: "ppt/slides/slide1.xml", seq: 1 }] };

async function merge(pkg: Pkg, rows: string[][], images?: Map<string, Uint8Array>) {
  const records = toRecordSet(rows);
  const plan = buildPlan(BLOCK, records, { runId: "r" });
  return runPlan(pkg, plan, records, { ...(images ? { images } : {}) });
}

/** The `<a:blipFill>` on a slide's first shape, if it has one. */
async function blipFill(pkg: Pkg, slide: string): Promise<Element | undefined> {
  const doc = await pkg.doc(slide);
  return elements(doc, A_NS, "blipFill")[0];
}

describe("a picture where the placeholder was", () => {
  it("fills the shape and takes the placeholder text away", async () => {
    /**
     * Both halves matter. Without the fill there is no picture; without
     * removing the text the file name prints ON TOP of the photo it asked for.
     */
    const pkg = await template("{{Photo|image}}", xfrm(200, 100));
    const out = await merge(pkg, [["Photo"], ["ada.png"]], new Map([["ada.png", WIDE]]));

    expect(out.images.placed).toBe(1);
    const slide = out.slides[0] as string;
    expect(await blipFill(pkg, slide), "no picture fill was written").toBeDefined();
    const text = elements(await pkg.doc(slide), A_NS, "t")
      .map((t) => t.textContent)
      .join("");
    expect(text, "the placeholder printed over the picture").not.toContain("{{Photo");
  });

  it("writes the media part, declares its type, and relates it from the slide", async () => {
    // Six things have to agree or PowerPoint reports the deck as damaged, and
    // it does that without saying which one was wrong.
    const pkg = await template("{{Photo|image}}", xfrm(200, 100));
    const out = await merge(pkg, [["Photo"], ["ada.png"]], new Map([["ada.png", WIDE]]));
    const slide = out.slides[0] as string;

    const media = pkg.partNames().filter((p) => p.startsWith("ppt/media/"));
    expect(media, "no media part was written").toHaveLength(1);
    expect(await pkg.bytes(media[0] as string), "the bytes were re-encoded").toEqual(WIDE);

    const types = await pkg.text("[Content_Types].xml");
    expect(types, "the png extension is not declared").toContain('Extension="png"');

    const fill = (await blipFill(pkg, slide)) as Element;
    const blip = elements(fill, A_NS, "blip")[0] as Element;
    const rId = blip.getAttribute("r:embed");
    expect(rId, "the fill points at no relationship").toBeTruthy();
    expect(await pkg.relTarget(slide, rId as string)).toBe(media[0]);
  });

  it("COVERS by default: crops the overflow rather than distorting", async () => {
    // A 64x32 image in a 200x100 shape. Both are 2:1, so nothing is trimmed —
    // and a 200x200 shape trims the sides.
    const square = await template("{{Photo|image}}", xfrm(200, 200));
    const out = await merge(square, [["Photo"], ["ada.png"]], new Map([["ada.png", WIDE]]));
    const fill = (await blipFill(square, out.slides[0] as string)) as Element;
    const srcRect = elements(fill, A_NS, "srcRect")[0] as Element;
    expect(srcRect, "cover wrote no srcRect").toBeDefined();
    expect(srcRect.getAttribute("l")).toBe("25000");
    expect(srcRect.getAttribute("r")).toBe("25000");
    expect(srcRect.getAttribute("t"), "cropped both axes").toBeNull();
  });

  it("CONTAINS on image-fit, because cropping a logo is somebody's trademark mangled", async () => {
    const pkg = await template("{{Logo|image-fit}}", xfrm(200, 200));
    const out = await merge(pkg, [["Logo"], ["mark.png"]], new Map([["mark.png", WIDE]]));
    const fill = (await blipFill(pkg, out.slides[0] as string)) as Element;
    expect(elements(fill, A_NS, "srcRect")[0], "contain cropped the image").toBeUndefined();
    const fillRect = elements(fill, A_NS, "fillRect")[0] as Element;
    // A wide image in a square box: bars above and below.
    expect(fillRect.getAttribute("t")).toBe("25000");
    expect(fillRect.getAttribute("b")).toBe("25000");
  });

  it("gives a shape with no geometry one, or the fill shows nothing at all", async () => {
    /**
     * A plain text box routinely has an empty `<p:spPr/>`. A picture fill on a
     * shape with no geometry renders NOTHING — no error, no placeholder, an
     * empty space where the photo should be, which is the worst way for this
     * to fail.
     */
    const pkg = await template("{{Photo|image}}");
    const out = await merge(pkg, [["Photo"], ["ada.png"]], new Map([["ada.png", WIDE]]));
    const doc = await pkg.doc(out.slides[0] as string);
    expect(elements(doc, A_NS, "prstGeom")[0], "no geometry, so no visible picture").toBeDefined();
  });

  it("gives one to a shape that is POSITIONED but has no geometry, in schema order", async () => {
    /**
     * The test above uses a shape whose `<p:spPr>` is empty. A shape with an
     * `<a:xfrm>` and no geometry is the commoner template — a text box somebody
     * placed — and it reaches a different branch: the predicate asks whether any
     * child is a geometry, and with a child present that answer can be wrong in
     * a way an empty `spPr` cannot show. `scripts/mutate-core.mjs` is what found
     * it: flipping that comparison left the suite green.
     *
     * The ORDER is asserted with it, and is the half no other test holds.
     * `CT_ShapeProperties` requires `xfrm` before the geometry, and the code
     * inserts after the `xfrm` precisely for that. Put the geometry first and
     * the file is schema-invalid — which PowerPoint answers by repairing the
     * deck and dropping what it chooses, the failure this whole suite exists to
     * keep out.
     */
    // An `<a:xfrm>` with something AFTER it, so the geometry has to be inserted
    // between them rather than appended. A fixture with only an xfrm reaches
    // the append branch instead, where a geometry inserted at the front would
    // still come out in the right order and the assertion would pass without
    // holding anything.
    const after = '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>';
    const pkg = await template("{{Photo|image}}", xfrm(200, 100) + after);
    const out = await merge(pkg, [["Photo"], ["ada.png"]], new Map([["ada.png", WIDE]]));
    const doc = await pkg.doc(out.slides[0] as string);

    const geom = elements(doc, A_NS, "prstGeom")[0];
    expect(geom, "a positioned shape with no geometry still shows no picture without one").toBeDefined();
    const order = Array.from((geom!.parentNode as Element).childNodes).map((n) => (n as Element).localName);
    expect(order.slice(0, 2)).toEqual(["xfrm", "prstGeom"]);
  });

  it("matches a photo whose cell the spreadsheet padded", async () => {
    // `resolveImage` trims the cell before matching it against the files, and
    // that `.trim()` is the whole rule. Removing it left the suite green —
    // so a Photo column exported with a trailing space would report every
    // picture missing, on a template that is correct, with the placeholder left
    // visible on every slide.
    const pkg = await template("{{Photo|image}}");
    const out = await merge(pkg, [["Photo"], ["  ada.png  "]], new Map([["ada.png", WIDE]]));

    expect(out.images.missing, "a padded cell is the same file").toEqual([]);
    expect(out.images.placed).toBe(1);
    expect(await blipFill(pkg, out.slides[0] as string)).toBeDefined();
  });

  it("STRETCHES and says so when the shape states no size", async () => {
    // Cover and contain need a ratio, and a shape that inherits its box from a
    // layout placeholder does not state one. Reported, because a stretched
    // photo reads as a broken image rather than a fact about the template.
    const pkg = await template("{{Photo|image}}");
    const out = await merge(pkg, [["Photo"], ["ada.png"]], new Map([["ada.png", WIDE]]));
    expect(out.images.stretched).toEqual(["Photo"]);
    const fill = (await blipFill(pkg, out.slides[0] as string)) as Element;
    expect(elements(fill, A_NS, "srcRect")[0], "cropped without a ratio to crop by").toBeUndefined();
  });
});

describe("a picture format this engine does not have", () => {
  /**
   * `image`, `image-fit` and `image-stretch` are the three modes, and anything
   * else fell through to `applyFormat`, whose documented answer for a format it
   * does not know is to return the cell unchanged. For a number that is the
   * right answer. For a picture it is not: the cell is a FILE NAME, so a single
   * transposed letter printed `ada.png` as text in the frame that was supposed
   * to hold the portrait, on every merged slide.
   *
   * `image-cover` is the likeliest of them, because the manual's own words for
   * what `image` does are "covers".
   *
   * The rule is about this engine's own namespace rather than a guess at what
   * the author meant: a format named `image`-something is an image format we do
   * not have, so the placeholder stays visible and the author sees their typo —
   * which is what a field with no column already does.
   */
  const MISSPELLINGS = ["image-cover", "images", "image-crop", "image-fill", "Image-Cover", " image-cover "];

  it("leaves the placeholder alone instead of printing the file name", () => {
    for (const spec of MISSPELLINGS) {
      const resolve = makeResolver({ Photo: "ada.png" });
      expect(resolve("Photo", spec), spec).toBeNull();
    }
  });

  it("still answers for a format that is not about pictures at all", () => {
    // The neighbouring rule this must not disturb: an unknown TEXT format
    // prints the cell, which is what the manual promises.
    const resolve = makeResolver({ Total: "1234.5" });
    expect(resolve("Total", "nubmer:2")).toBe("1234.5");
    expect(resolve("Total", "imagination")).toBe("1234.5");
  });

  it("and the three real ones still place a picture", () => {
    for (const spec of ["image", "image-fit", "image-stretch"]) {
      expect(imageMode(spec), spec).toBeDefined();
    }
  });
});

describe("what a merge does with an image it cannot get", () => {
  it("leaves the placeholder VISIBLE when no file was supplied", async () => {
    // The rule text fields already follow, for the same reason: a blank frame
    // looks finished and is not.
    const pkg = await template("{{Photo|image}}", xfrm(200, 100));
    const out = await merge(pkg, [["Photo"], ["ada.png"]]);
    expect(out.images.placed).toBe(0);
    expect(out.images.missing).toEqual(["Photo"]);
    const text = elements(await pkg.doc(out.slides[0] as string), A_NS, "t")
      .map((t) => t.textContent)
      .join("");
    expect(text).toContain("{{Photo|image}}");
  });

  it("reports a file that is not an image, and does not embed it", async () => {
    const pkg = await template("{{Photo|image}}", xfrm(200, 100));
    const junk = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]);
    const out = await merge(pkg, [["Photo"], ["ada.png"]], new Map([["ada.png", junk]]));
    expect(out.images.unreadable).toEqual(["Photo"]);
    expect(
      pkg.partNames().filter((p) => p.startsWith("ppt/media/")),
      "junk reached the deck",
    ).toHaveLength(0);
  });

  it("matches a cell that carries a PATH, and ignores case", async () => {
    // `Photos\\ada.PNG` in a spreadsheet and `ada.png` from a file picker are
    // the same picture, and a user should not have to know that.
    for (const cell of ["Photos\\ada.PNG", "photos/Ada.png", "ADA.PNG"]) {
      const pkg = await template("{{Photo|image}}", xfrm(200, 100));
      const out = await merge(pkg, [["Photo"], [cell]], new Map([["ada.png", WIDE]]));
      expect(out.images.placed, cell).toBe(1);
    }
  });
});

describe("two hundred rows of the same logo", () => {
  it("stores the bytes ONCE, however many rows use them", async () => {
    /**
     * Not a nicety. The package goes to the host as a single base64 string, so
     * 240 copies of a 200 KB logo is 48 MB become 64 MB across the wire, to a
     * host that already stalls on ordinary work.
     */
    const pkg = await template("{{Logo|image}}", xfrm(200, 100));
    const rows = [["Logo"], ...Array.from({ length: 20 }, () => ["mark.png"])];
    const out = await merge(pkg, rows, new Map([["mark.png", WIDE]]));

    expect(out.slides).toHaveLength(20);
    expect(out.images.placed).toBe(20);
    expect(pkg.partNames().filter((p) => p.startsWith("ppt/media/"))).toHaveLength(1);
  });

  it("keeps two DIFFERENT pictures apart", async () => {
    const pkg = await template("{{Photo|image}}", xfrm(200, 100));
    const out = await merge(
      pkg,
      [["Photo"], ["ada.png"], ["grace.jpg"]],
      new Map([
        ["ada.png", WIDE],
        ["grace.jpg", TALL],
      ]),
    );
    expect(out.images.placed).toBe(2);
    const media = pkg
      .partNames()
      .filter((p) => p.startsWith("ppt/media/"))
      .sort();
    expect(media).toHaveLength(2);
    // Named by their own format, not by the field's.
    expect(media.some((m) => m.endsWith(".png"))).toBe(true);
    expect(media.some((m) => m.endsWith(".jpeg"))).toBe(true);
  });
});

describe("the package a picture merge writes", () => {
  it("opens as a zip whose every relationship resolves", async () => {
    /**
     * The end-to-end check. PowerPoint reports a package with a dangling
     * relationship as a damaged file and does not say which part it could not
     * find, so this walks every `.rels` in the finished bytes and asserts the
     * target is really there.
     */
    const pkg = await template("{{Photo|image}}", xfrm(200, 100));
    await merge(
      pkg,
      [["Photo"], ["ada.png"], ["grace.jpg"]],
      new Map([
        ["ada.png", WIDE],
        ["grace.jpg", TALL],
      ]),
    );

    const zip = await JSZip.loadAsync(await pkg.toBytes());
    const names = new Set(Object.keys(zip.files));
    let checked = 0;
    for (const path of [...names].filter((n) => n.endsWith(".rels"))) {
      const owner = path.replace("_rels/", "").replace(/\.rels$/, "");
      const xml = await (zip.file(path) as JSZip.JSZipObject).async("string");
      for (const match of xml.matchAll(/Target="([^"]+)"/g)) {
        const target = match[1] ?? "";
        if (target.startsWith("http") || target === "") continue;
        const base = owner.split("/").slice(0, -1);
        const to = target.replace(/^\//, "").split("/");
        const resolved = target.startsWith("/") ? to.join("/") : resolveRel(base, to);
        expect(names.has(resolved), `${path} points at a missing ${resolved}`).toBe(true);
        checked++;
      }
    }
    // The walk is only worth anything if it actually walked something.
    expect(checked, "no relationships were checked at all").toBeGreaterThan(5);
  });

  it("declares a content type for every part it wrote", async () => {
    const pkg = await template("{{Photo|image}}", xfrm(200, 100));
    await merge(pkg, [["Photo"], ["ada.png"]], new Map([["ada.png", WIDE]]));
    const types = await pkg.text("[Content_Types].xml");
    for (const part of pkg.partNames().filter((p) => p.startsWith("ppt/media/"))) {
      const ext = part.slice(part.lastIndexOf(".") + 1);
      expect(types, `${part} has no declared type`).toContain(`Extension="${ext}"`);
    }
  });

  it("adds ONE default per extension, however many pictures use it", async () => {
    // A second Default for the same extension is schema-invalid, and a merge
    // that embeds two hundred photos would otherwise add two hundred of them.
    const pkg = await template("{{Photo|image}}", xfrm(200, 100));
    await merge(
      pkg,
      [["Photo"], ["a.png"], ["b.png"]],
      new Map([
        ["a.png", WIDE],
        ["b.png", TALL],
      ]),
    );
    const types = await pkg.text("[Content_Types].xml");
    expect(types.match(/Extension="png"/g) ?? []).toHaveLength(1);
  });
});

function resolveRel(base: string[], to: string[]): string {
  const parts = [...base];
  for (const seg of to) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

describe("what the picture pass must not take with it", () => {
  /**
   * The pass placed its picture and then blanked EVERY text node in the
   * paragraph. The placeholder went, which was the intent, and so did whatever
   * else shared the paragraph: a caption, or another field.
   *
   * The second is the one worth a test of its own. A caption disappearing is
   * visible — somebody notices the slide is bare. A merged VALUE disappearing
   * leaves a slide that looks finished and is missing the thing the merge was
   * for, on every copy, with nothing in any count saying so.
   */
  async function mergeShape(paras: string[][], rows: string[][], images: Map<string, Uint8Array>) {
    const pkg = await Pkg.open(
      await makeDeck([{ paragraphs: paras, box: xfrm(200, 100) }, { paragraphs: [["after"]] }]),
    );
    const records = toRecordSet(rows);
    const out = await runPlan(pkg, buildPlan(BLOCK, records, { runId: "r" }), records, { images });
    const doc = await pkg.doc(out.slides[0] as string);
    return {
      images: out.images,
      text: elements(doc, A_NS, "t")
        .map((t) => t.textContent)
        .join(""),
      fills: elements(doc, A_NS, "blipFill").length,
    };
  }

  const one = () => new Map([["ada.png", WIDE]]);
  const two = () =>
    new Map([
      ["ada.png", WIDE],
      ["bo.jpg", TALL],
    ]);

  it("keeps a caption written beside the placeholder", async () => {
    const r = await mergeShape([["Photo: ", "{{Photo|image}}"]], [["Photo"], ["ada.png"]], one());
    expect(r.text, "the caption was blanked with the placeholder").toBe("Photo: ");
    expect(r.fills, "no picture was placed").toBe(1);
  });

  it("keeps a merged value written beside the placeholder", async () => {
    const r = await mergeShape(
      [["{{Name}} ", "{{Photo|image}}"]],
      [
        ["Name", "Photo"],
        ["Ada", "ada.png"],
      ],
      one(),
    );
    expect(r.text.trim(), "the row's own value was blanked with the placeholder").toBe("Ada");
    expect(r.fills).toBe(1);
  });

  it("fills a shape once and says so when a second field wanted it", async () => {
    /**
     * A shape has ONE fill. The second field used to overwrite the first,
     * count itself into `placed`, and leave a media part and a relationship
     * behind for a picture that is not on the slide — a count saying two where
     * the deck shows one.
     */
    const r = await mergeShape(
      [["{{A|image}} {{B|image}}"]],
      [
        ["A", "B"],
        ["ada.png", "bo.jpg"],
      ],
      two(),
    );
    expect(r.fills).toBe(1);
    expect(r.images.placed, "counted a picture that is not on the slide").toBe(1);
    expect(r.images.crowded, "the second field was dropped in silence").toEqual(["B"]);
    // Left standing, so the author sees which one could not be drawn.
    expect(r.text).toContain("{{B|image}}");
  });

  it("counts the same way when the second field is in its own paragraph", async () => {
    const r = await mergeShape(
      [["{{A|image}}"], ["{{B|image}}"]],
      [
        ["A", "B"],
        ["ada.png", "bo.jpg"],
      ],
      two(),
    );
    expect(r.fills).toBe(1);
    expect(r.images.placed).toBe(1);
    expect(r.images.crowded).toEqual(["B"]);
  });

  it("reports an image field in a table cell rather than skipping it", async () => {
    /**
     * The file said this happened long before it could. The walk started at
     * `<p:sp>`, so a table's paragraph was never visited and the check for it
     * sat below a loop that could not reach it — dead code under a sentence
     * promising the behaviour. The field got no picture and no mention.
     */
    const table =
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Table"/>' +
      "<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>" +
      '<p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></p:xfrm>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
      '<a:tbl><a:tr h="1000"><a:tc><a:txBody><a:bodyPr/>' +
      "<a:p><a:r><a:t>{{Photo|image}}</a:t></a:r></a:p>" +
      "</a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>";

    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["plain"]] }, { paragraphs: [["after"]] }]));
    const path = "ppt/slides/slide1.xml";
    const xml = new TextDecoder().decode(await pkg.bytes(path));
    pkg.setBytes(path, new TextEncoder().encode(xml.replace("</p:spTree>", table + "</p:spTree>")));

    const records = toRecordSet([["Photo"], ["ada.png"]]);
    const out = await runPlan(pkg, buildPlan(BLOCK, records, { runId: "r" }), records, {
      images: one(),
    });

    expect(out.images.missing, "a field with no shape was skipped in silence").toEqual(["Photo"]);
    expect(out.images.placed).toBe(0);
  });
});
