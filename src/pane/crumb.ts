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
  /**
   * WHICH document the run was against. The store is not scoped to one.
   *
   * `localStorage` belongs to the add-in's ORIGIN, and every deck the user opens
   * shares it. Without this the crumb from one deck's interrupted run was read
   * while another was open: the real-host round of 2026-08-30 opened a fresh,
   * never-merged deck and was told "a merge from 2026-08-30 added 6 slide(s)".
   *
   * That instance was only alarming — `sweepPlan` saw a deck that had not grown
   * and refused, so nothing was offered. It is not safe in general, because
   * every clamp in `sweepPlan` compares SIZES and none of them compares
   * identity: a stranger deck holding exactly `deckAtStart + added` slides
   * satisfies all of them and yields a plan to delete slides the run never
   * created. The missing question was never "how big" but "which deck".
   */
  doc: string;
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

/**
 * Forget it — the run finished, or the user took the slides back.
 *
 * `here` is the open document's key, and it is asked for the same reason
 * `readCrumb` asks: the store is ONE key for every deck the user opens, because
 * `localStorage` belongs to the add-in's origin. Without it, a run in one deck
 * deleted the recovery record of another — the read side refused a stranger's
 * crumb at length and the write side deleted it without looking, which is the
 * asymmetry that makes a careful check worthless.
 *
 * A crumb this build cannot identify — no `doc`, an older build's shape, or
 * anything unparseable — IS cleared. Refusing what cannot be matched would make
 * the key unreclaimable: nothing could ever remove it and every future run
 * would step around it. The rule is "another deck's crumb is safe", not "only a
 * crumb I can name may go".
 */
export function clearCrumb(here: string): void {
  const s = store();
  if (!s) return;
  try {
    const raw = s.getItem(KEY);
    // Read through `readCrumb` rather than re-implementing the shape check, so
    // "a crumb this build understands" cannot come to mean two different things
    // in the two functions that ask it.
    if (raw && readCrumb(here) === undefined && belongsToAnotherDeck(raw, here)) return;
    s.removeItem(KEY);
  } catch {
    /* nothing to be done, and nothing worth failing over */
  }
}

/** Whether a stored record names a deck, and names a different one from `here`. */
function belongsToAnotherDeck(raw: string, here: string): boolean {
  try {
    const parsed: unknown = JSON.parse(raw);
    if (typeof parsed !== "object" || parsed === null) return false;
    const doc = (parsed as Partial<Crumb>).doc;
    return typeof doc === "string" && doc !== "" && doc !== here;
  } catch {
    return false;
  }
}

/**
 * What the last run left, when it left anything AND it belongs to this deck.
 *
 * Returns undefined for an absent, unreadable, or unrecognisable crumb —
 * including one from an older build. The undo path acts on these numbers, so
 * "I could not read it" and "the deck was this big" must never be confused.
 *
 * `here` is the open document's key, handed in rather than read. Only
 * `src/pane/main.ts` may touch Office.js: the moment this file asked the host
 * anything, the labels a user presses would stop being checkable without a
 * PowerPoint, and the pane is exactly where a wrong label is the thing that
 * gets pressed. Pass "" for a host that will not name the document — it is
 * treated as "cannot prove", never as a match.
 */
export function readCrumb(here: string): Crumb | undefined {
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
    // WHICH deck, asked before the numbers are handed to anything that deletes.
    //
    // Refused unless both keys are known and identical. A crumb from an older
    // build carries no `doc` at all, and an unreadable host answers "", and
    // neither can be told apart from a match by guessing — so both are refused.
    //
    // The cost is a real one: on a host that will not name the document, the
    // dead-tab recovery this file exists for stops being offered. That is the
    // direction this project already takes when identity cannot be proven —
    // `sweepPlan` leaves slides behind rather than delete one it cannot show it
    // created — and the alternative is offering to delete six slides of
    // whatever deck happens to be the right size.
    if (!here || typeof c.doc !== "string" || c.doc === "" || c.doc !== here) return undefined;
    return {
      kind: "ssf-merge-run",
      deckAtStart: c.deckAtStart as number,
      added: c.added as number,
      runId: typeof c.runId === "string" ? c.runId : "unknown",
      startedAt: typeof c.startedAt === "string" ? c.startedAt : "unknown",
      doc: c.doc,
    };
  } catch {
    return undefined;
  }
}
