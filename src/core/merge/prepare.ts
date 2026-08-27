/**
 * Turn what the pane knows into what the engine needs, and say no clearly.
 *
 * The pane speaks in SLIDE NUMBERS — the ones in the thumbnail rail, which is
 * the only numbering a user can see. The engine speaks in package paths. This
 * is the translation, and it is pure so that every refusal can be checked
 * without a PowerPoint: the alternative is a merge that starts, clones the
 * wrong slides, and is discovered by looking at the output.
 */
import type { Pkg } from "../pptx/pkg.js";
import type { Block, BlockSlide } from "./plan.js";
import { notesPathFor } from "../pptx/clone.js";
import { fieldsIn } from "./text.js";

export interface BlockRequest {
  /** First slide of the template, 1-based, as the thumbnail rail shows it. */
  from: number;
  /** Last slide, 1-based, inclusive. */
  to: number;
  /** Where the block starts INSIDE the package that was read back. */
  offsetInPackage: number;
  /** Conditions the pane collected, keyed by slide number. */
  conditions?: Record<number, string>;
}

export type Prepared = { ok: true; block: Block; fields: string[]; unmergeable: string[] } | { ok: false; why: string };

/**
 * Parts that hold text this engine does not merge.
 *
 * A chart's labels live in `ppt/charts/chartN.xml` with an embedded workbook
 * behind them, and SmartArt's live in `ppt/diagrams/dataN.xml`. Neither is a
 * `<a:p>` on the slide, so `mergeDocument` never reaches them and `fieldsIn`
 * never reports them.
 *
 * Not merging them is a stated limit. Not SAYING so is the defect: the author
 * puts `{{Name}}` in a chart title, the pane counts the placeholders it can see
 * and says nothing about that one, and 240 slides ship with the braces on them.
 */
const UNMERGED_PARTS = /^ppt\/(charts\/chart\d+|diagrams\/data\d+)\.xml$/;

/**
 * Build the block, or refuse with a sentence the pane can show as it stands.
 *
 * Every refusal here is a thing the user can act on. "Something went wrong" is
 * the failure this function exists to avoid: a merge is expensive to undo and
 * the user is the only one who knows which slides they meant.
 */
export async function prepareBlock(pkg: Pkg, req: BlockRequest, runId: string): Promise<Prepared> {
  if (!Number.isInteger(req.from) || !Number.isInteger(req.to)) {
    return { ok: false, why: "The template block has to be whole slide numbers." };
  }
  if (req.from < 1) return { ok: false, why: "Slides are numbered from 1." };
  if (req.to < req.from) {
    return { ok: false, why: `The block ends before it starts: slide ${req.from} to ${req.to}.` };
  }

  const paths = await pkg.slidePaths();
  // The package read back is either the whole deck or just the block, and the
  // two number their slides differently. `offsetInPackage` is what reconciles
  // them; getting it wrong merges the wrong slides and the output looks
  // deliberate, which is why `templateOffset` is its own tested function.
  const start = req.offsetInPackage;
  const count = req.to - req.from + 1;
  if (start < 0 || start + count > paths.length) {
    return {
      ok: false,
      why: `The template block is slides ${req.from} to ${req.to}, and the deck that came back has ${paths.length}.`,
    };
  }

  const slides: BlockSlide[] = [];
  const fields: string[] = [];
  const unmergeable: string[] = [];
  for (let i = 0; i < count; i++) {
    const path = paths[start + i];
    if (!path) return { ok: false, why: `Slide ${req.from + i} is not in the deck that came back.` };
    // The slide, and its speaker notes.
    //
    // `runPlan` merges the notes page — a template whose notes read "Call
    // {{Name}} afterwards" otherwise ships that verbatim on every handout — but
    // this scan only ever read the slide. So a block whose placeholders live in
    // the notes was refused with "no placeholders, so every copy would be
    // identical", about a merge that would have filled them.
    //
    // The mirror of the chart case below, and the worse direction: there the
    // pane reported fields it cannot merge, here it hid fields it can and
    // blocked the merge on the strength of it.
    const notes = await notesPathFor(pkg, path);
    const own = [...fieldsIn(await pkg.doc(path)), ...(notes ? fieldsIn(await pkg.doc(notes)) : [])];
    for (const f of own) if (!fields.includes(f)) fields.push(f);
    // Fields the author placed somewhere this engine does not reach. Read from
    // the parts THIS slide relates to, never from the package at large: on the
    // route where the template came back as the whole deck, a chart on slide 40
    // is not this block's problem and naming it would send the user hunting.
    for (const part of await pkg.relatedParts(path)) {
      if (!UNMERGED_PARTS.test(part) || !pkg.has(part)) continue;
      // `fieldsIn`, not a regex over the raw XML. Chart and SmartArt text is
      // DrawingML — the same `<a:p>` and `<a:t>` the slide uses — so the same
      // reader finds a placeholder PowerPoint has split across runs, which is
      // the ordinary state of one after an edit. A regex over the markup would
      // miss exactly those and report the tidy ones.
      for (const name of fieldsIn(await pkg.doc(part))) {
        if (!unmergeable.includes(name)) unmergeable.push(name);
      }
    }
    const condition = req.conditions?.[req.from + i];
    slides.push({ path, seq: i + 1, fields: own, ...(condition ? { condition } : {}) });
  }

  if (fields.length === 0) {
    // Not an error the engine can see: a merge with no placeholders produces N
    // identical copies, which is never what anybody meant and is expensive to
    // undo once it is in the deck.
    const where = count === 1 ? `Slide ${req.from}` : `Slides ${req.from} to ${req.to}`;
    // The block whose ONLY placeholders are in a chart. "No placeholders" is
    // true and useless: the author placed one, is looking at it, and is being
    // told it is not there. Which is the whole complaint this pass exists for,
    // arriving on the one path where it reads as the engine being wrong.
    if (unmergeable.length > 0) {
      return {
        ok: false,
        why:
          `${where} has ${unmergeable.length === 1 ? "a placeholder" : "placeholders"} only inside a chart or ` +
          `SmartArt (${unmergeable.join(", ")}), which SSF Merge cannot fill. Move the text onto the slide ` +
          `itself, or mark a block that has a placeholder on the slide.`,
      };
    }
    return { ok: false, why: `${where} has no placeholders, so every copy would be identical.` };
  }

  return { ok: true, block: { id: runId, slides }, fields, unmergeable };
}
