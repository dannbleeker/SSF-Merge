/**
 * The two numbers an undo cannot be done without, kept where a dead tab cannot
 * take them.
 *
 * This is not a diagnostic. `undoInsert` is positional and clamped against the
 * deck's size BEFORE the run inserted anything, and those numbers live in a
 * module variable — so a tab that dies during a merge leaves the user's deck
 * holding 720 new slides and the pane with no way to take them back. The slides
 * are there; the only thing missing is two integers.
 *
 * A sibling project's answer to the same shape is 445 lines: two storage slots
 * with promotion, a sequence counter, a debounced flush, halve-and-retry on a
 * full store, a findings channel with byte caps. It needs all of that because
 * it is preserving a 300-entry narrative of a fifteen-minute run. This is
 * preserving `{deckAtStart, added, runId}`, so it is one key and one write.
 *
 * Three rules, all of them the sibling's and all of them cheap to obey:
 *
 * - **Never throw.** Storage can be absent, disabled by policy — a task pane is
 *   a third-party frame — or full, and none of that is a reason for a merge to
 *   fail. Every path here swallows.
 * - **Probe with a READ.** Asking "is there a store" with a write conflates it
 *   with "is there room in it", and a full store answers no to the second while
 *   the record we need to read back sits in it needing no quota at all.
 * - **Validate on the way out.** A crumb written by an older build, or half
 *   written by a crash mid-write, must not reach the undo path as if it were
 *   sound.
 */

const KEY = "ssf-merge.run.v1";

/** What a run left behind, if it did not get to finish. */
export interface Crumb {
  /** Names the shape for whatever reads it, and dates it for whoever changes it. */
  kind: "ssf-merge-run";
  /** The deck's size before the insert. The floor every sweep is clamped to. */
  deckAtStart: number;
  /**
   * What the run believes it added. ZERO until the insert answers.
   *
   * Zero is a real state and not a broken record: the crumb is written BEFORE
   * the insert precisely so a tab that dies during it leaves something behind.
   * A reader must not sweep on it — `sweepPlan` refuses a count of zero, and
   * deriving one from the deck's growth would delete whatever has been appended
   * since, which is the inference this whole file exists to refuse. What it CAN
   * do is say a run did not finish and how big the deck was before it.
   */
  added: number;
  runId: string;
  startedAt: string;
}

/** The store, or null where there is not one. Never throws. */
function store(): Storage | null {
  try {
    const s = globalThis.localStorage;
    // Touched with a READ. A store that exists and throws on access is a real
    // configuration, and it throws at USE rather than at lookup.
    s.getItem(KEY);
    return s;
  } catch {
    return null;
  }
}

/** Remember what an undo would need, before the insert that makes it necessary. */
export function dropCrumb(c: Omit<Crumb, "kind" | "startedAt">): void {
  const s = store();
  if (!s) return;
  try {
    const crumb: Crumb = { kind: "ssf-merge-run", startedAt: new Date().toISOString(), ...c };
    s.setItem(KEY, JSON.stringify(crumb));
  } catch {
    /* a merge does not fail because a browser would not remember something */
  }
}

/** Forget it — the run finished, or the user took the slides back. */
export function clearCrumb(): void {
  const s = store();
  if (!s) return;
  try {
    s.removeItem(KEY);
  } catch {
    /* nothing to be done, and nothing worth failing over */
  }
}

/**
 * What the last run left, when it left anything.
 *
 * Returns undefined for an absent, unreadable, or unrecognisable crumb —
 * including one from an older build. The undo path acts on these numbers, so
 * "I could not read it" and "the deck was this big" must never be confused.
 */
export function readCrumb(): Crumb | undefined {
  const s = store();
  if (!s) return undefined;
  try {
    const raw = s.getItem(KEY);
    if (!raw) return undefined;
    const parsed: unknown = JSON.parse(raw);
    if (typeof parsed !== "object" || parsed === null) return undefined;
    const c = parsed as Partial<Crumb>;
    if (c.kind !== "ssf-merge-run") return undefined;
    // Both numbers, or neither. A crumb carrying one of them cannot clamp a
    // sweep, and a sweep that cannot be clamped is not offered.
    if (!Number.isInteger(c.deckAtStart) || !Number.isInteger(c.added)) return undefined;
    // Zero added is ALLOWED, and refusing it made this file write-only in the
    // one window it was built for. `merge()` writes `{added: 0}` before the
    // insert — that is the whole point, because a tab that dies during the
    // insert never gets to write the real number — and this line then threw it
    // away on the way back in. A crumb was therefore readable only after the
    // run it was insurance against had already succeeded.
    //
    // Zero still authorises nothing: `sweepPlan` refuses a count of zero, so
    // the caller gets a sentence rather than a button. Negative is still a
    // broken record.
    if ((c.deckAtStart ?? -1) < 0 || (c.added ?? -1) < 0) return undefined;
    return {
      kind: "ssf-merge-run",
      deckAtStart: c.deckAtStart as number,
      added: c.added as number,
      runId: typeof c.runId === "string" ? c.runId : "unknown",
      startedAt: typeof c.startedAt === "string" ? c.startedAt : "unknown",
    };
  } catch {
    return undefined;
  }
}
