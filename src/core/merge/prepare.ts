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

export type Prepared = { ok: true; block: Block; fields: string[] } | { ok: false; why: string };

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
  for (let i = 0; i < count; i++) {
    const path = paths[start + i];
    if (!path) return { ok: false, why: `Slide ${req.from + i} is not in the deck that came back.` };
    const own = fieldsIn(await pkg.doc(path));
    for (const f of own) if (!fields.includes(f)) fields.push(f);
    const condition = req.conditions?.[req.from + i];
    slides.push({ path, seq: i + 1, fields: own, ...(condition ? { condition } : {}) });
  }

  if (fields.length === 0) {
    // Not an error the engine can see: a merge with no placeholders produces N
    // identical copies, which is never what anybody meant and is expensive to
    // undo once it is in the deck.
    return {
      ok: false,
      why:
        count === 1
          ? `Slide ${req.from} has no placeholders, so every copy would be identical.`
          : `Slides ${req.from} to ${req.to} have no placeholders, so every copy would be identical.`,
    };
  }

  return { ok: true, block: { id: runId, slides }, fields };
}
