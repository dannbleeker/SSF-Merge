/**
 * Put a picture into a shape that already exists.
 *
 * The shape is FILLED, never replaced. A `<p:pic>` swapped in for the `<p:sp>`
 * would have to carry the geometry the shape had — and a shape that takes its
 * position from a layout placeholder does not state one, so the swap would have
 * to invent numbers or lose the placement. Filling keeps the position, the size,
 * the rotation, the group it sits in and the placeholder it inherits from, and
 * changes exactly one thing.
 *
 * It is also the answer the host measurement pointed at. `ShapeFill.setImage`
 * STRETCHES on PowerPoint for the web — measured 2026-08-28, a square card into
 * a 2:1 box comes out a wide ellipse — so anything that went through the API
 * would have to letterbox the image itself before sending. In the package the
 * fill mode is ours to write, and `<a:blipFill>` states it directly.
 */
import { A_NS, P_NS, R_NS, child } from "../pptx/xml.js";

/**
 * Every ELEMENT child, whatever its name.
 *
 * `children` in `xml.ts` filters by namespace and local name, which is what
 * every other caller wants. Ordering a sequence needs the whole list — a fill
 * has to go after the geometry whatever the geometry is called, and there are
 * two spellings of it.
 */
function allChildren(parent: Element): Element[] {
  const out: Element[] = [];
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1) out.push(n as Element);
  }
  return out;
}
import { coverSrcRect, containFillRect, NO_INSET, type FillMode, type Inset } from "./fill.js";

/** Where a shape's own size comes from, when it states one. */
export interface Box {
  w: number;
  h: number;
}

/**
 * A shape's size in EMU, or undefined when it inherits one.
 *
 * A shape with no `<a:xfrm>` takes its box from the layout placeholder it
 * points at, and following that would mean resolving the layout, then the
 * master, then their placeholder types. Undefined is the honest answer here and
 * the caller degrades: without a box there is no ratio, so cover and contain
 * cannot be computed and the fill falls back to a plain stretch.
 */
export function shapeBox(sp: Element): Box | undefined {
  const spPr = child(sp, P_NS, "spPr");
  const xfrm = spPr ? child(spPr, A_NS, "xfrm") : undefined;
  const ext = xfrm ? child(xfrm, A_NS, "ext") : undefined;
  if (!ext) return undefined;
  let w = Number(ext.getAttribute("cx") ?? 0);
  let h = Number(ext.getAttribute("cy") ?? 0);
  if (!(w > 0 && h > 0)) return undefined;
  // Scaled by every GROUP this shape is inside.
  //
  // A child of a `<p:grpSp>` states its size in the group's CHILD coordinate
  // space, and the group scales that space by `ext ÷ chExt`. PowerPoint writes
  // the two equal for a new group and leaves `chExt` alone when the user drags
  // the group's handles — so every group anybody has resized has a scale
  // factor, and the box this used to answer was not the box the picture is seen
  // in. `cover`, whose whole job is not to distort, then computed its crop for
  // the wrong ratio: a group stretched 4:1 horizontally cropped the SIDES off a
  // 2:1 photo and had the group stretch the remaining square, on the one run
  // that asked not to be squashed. Nothing reported it, because a crop was
  // computed and written.
  //
  // Only the RATIO is used downstream, so the factors are applied to `w` and
  // `h` directly and no EMU rounding enters. Groups nest, so this accumulates
  // rather than looking one level up; a level whose numbers are missing or
  // non-positive is skipped, which leaves the box as it stands rather than
  // discarding a scale it cannot read.
  for (let at = sp.parentNode; at !== null; at = at.parentNode) {
    if (at.nodeType !== 1) continue;
    const node = at as Element;
    if (node.namespaceURI !== P_NS || node.localName !== "grpSp") continue;
    const props = child(node, P_NS, "grpSpPr");
    const groupXfrm = props ? child(props, A_NS, "xfrm") : undefined;
    if (!groupXfrm) continue;
    const outer = child(groupXfrm, A_NS, "ext");
    const inner = child(groupXfrm, A_NS, "chExt");
    if (!outer || !inner) continue;
    const [ox, oy] = [Number(outer.getAttribute("cx") ?? 0), Number(outer.getAttribute("cy") ?? 0)];
    const [ix, iy] = [Number(inner.getAttribute("cx") ?? 0), Number(inner.getAttribute("cy") ?? 0)];
    if (ox > 0 && ix > 0) w *= ox / ix;
    if (oy > 0 && iy > 0) h *= oy / iy;
  }
  return w > 0 && h > 0 ? { w, h } : undefined;
}

/**
 * The mode to use, given what is known.
 *
 * Named rather than inlined because the degradation is the interesting part: a
 * shape that states no box has no ratio, so the mode the author asked for
 * cannot be computed and the picture stretches. The caller reports that; it
 * must not be silent, because a stretched photo looks like a bug in the image
 * rather than a fact about the template.
 */
export function modeFor(asked: FillMode, box: Box | undefined, image: Box | undefined): FillMode {
  if (asked === "stretch") return "stretch";
  return box && image ? asked : "stretch";
}

function insetFor(mode: FillMode, box: Box | undefined, image: Box | undefined): Inset {
  if (!box || !image) return NO_INSET;
  if (mode === "cover") return coverSrcRect(box, image);
  if (mode === "contain") return containFillRect(box, image);
  return NO_INSET;
}

function insetAttrs(node: Element, inset: Inset): void {
  // Only the sides that are actually inset. A `<a:srcRect/>` with four zeroes
  // is legal and means the same as an empty one, but the empty one is what
  // PowerPoint writes and the smaller diff is easier to read in a package
  // somebody is debugging.
  for (const [name, value] of [
    ["l", inset.l],
    ["t", inset.t],
    ["r", inset.r],
    ["b", inset.b],
  ] as const) {
    if (value > 0) node.setAttribute(name, String(value));
  }
}

/**
 * `<a:spPr>`'s children are an ORDERED sequence, and a fill in the wrong place
 * is a part PowerPoint rejects.
 *
 * `CT_ShapeProperties` is `a:xfrm?, <geometry>?, <fill>?, a:ln?, <effects>?,
 * a:scene3d?, a:sp3d?, a:extLst?`. So a blipFill goes after the geometry and
 * before the line, and every other fill has to come out first — a shape can
 * hold one.
 */
const FILL_KINDS = ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];
const BEFORE_FILL = ["xfrm", "custGeom", "prstGeom"];

function insertFill(spPr: Element, fill: Element): void {
  for (const existing of allChildren(spPr).filter((c) => FILL_KINDS.includes(c.localName))) {
    spPr.removeChild(existing);
  }
  const after = allChildren(spPr)
    .filter((c) => BEFORE_FILL.includes(c.localName))
    .pop();
  if (after?.nextSibling) spPr.insertBefore(fill, after.nextSibling);
  else if (after) spPr.appendChild(fill);
  else spPr.insertBefore(fill, spPr.firstChild);
}

/**
 * Make sure the shape has a geometry, because a fill needs one to show.
 *
 * A plain text box routinely has an empty `<p:spPr/>`: its geometry is a
 * rectangle by default as far as the text is concerned, and nothing declares
 * it. A picture fill on a shape with no geometry renders nothing at all — no
 * error, no placeholder, an empty space where the photo should be, which is the
 * worst way for this to fail.
 */
function ensureGeometry(doc: Document, spPr: Element): void {
  if (allChildren(spPr).some((c) => c.localName === "prstGeom" || c.localName === "custGeom")) return;
  const geom = doc.createElementNS(A_NS, "a:prstGeom");
  geom.setAttribute("prst", "rect");
  geom.appendChild(doc.createElementNS(A_NS, "a:avLst"));
  const xfrm = child(spPr, A_NS, "xfrm");
  if (xfrm?.nextSibling) spPr.insertBefore(geom, xfrm.nextSibling);
  else if (xfrm) spPr.appendChild(geom);
  else spPr.insertBefore(geom, spPr.firstChild);
}

export interface PlaceResult {
  /** The mode actually used, which is `stretch` when a ratio could not be had. */
  mode: FillMode;
}

/**
 * Fill one shape with the media part `rId` points at.
 *
 * `image` is the picture's pixel size and may be undefined when the bytes could
 * not be read; `asked` is the mode the field requested. Together with the
 * shape's own box they decide the fill, and the result says what was done so
 * the run can report a stretch nobody asked for.
 */
export function fillShapeWithImage(
  doc: Document,
  sp: Element,
  rId: string,
  asked: FillMode,
  image: Box | undefined,
): PlaceResult {
  let spPr = child(sp, P_NS, "spPr");
  if (!spPr) {
    spPr = doc.createElementNS(P_NS, "p:spPr");
    // After the non-visual properties, which `CT_Shape` requires first.
    const nv = child(sp, P_NS, "nvSpPr");
    if (nv?.nextSibling) sp.insertBefore(spPr, nv.nextSibling);
    else sp.appendChild(spPr);
  }
  ensureGeometry(doc, spPr);

  const box = shapeBox(sp);
  const mode = modeFor(asked, box, image);
  const inset = insetFor(mode, box, image);

  const fill = doc.createElementNS(A_NS, "a:blipFill");
  const blip = doc.createElementNS(A_NS, "a:blip");
  blip.setAttributeNS(R_NS, "r:embed", rId);
  fill.appendChild(blip);

  if (mode === "cover") {
    const srcRect = doc.createElementNS(A_NS, "a:srcRect");
    insetAttrs(srcRect, inset);
    fill.appendChild(srcRect);
  }
  const stretch = doc.createElementNS(A_NS, "a:stretch");
  const fillRect = doc.createElementNS(A_NS, "a:fillRect");
  if (mode === "contain") insetAttrs(fillRect, inset);
  stretch.appendChild(fillRect);
  fill.appendChild(stretch);

  insertFill(spPr, fill);
  return { mode };
}

/** The shape a node sits inside, or undefined. Used to find the box a field is in. */
export function shapeOf(node: Node): Element | undefined {
  let at: Node | null = node;
  while (at) {
    if (at.nodeType === 1 && (at as Element).localName === "sp" && (at as Element).namespaceURI === P_NS) {
      return at as Element;
    }
    at = at.parentNode;
  }
  return undefined;
}
