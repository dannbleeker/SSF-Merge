import { trace } from "../core/trace.js";

/**
 * Every host call is bounded.
 *
 * On PowerPoint for the web a call that stops answering does not come back:
 * across a sibling project's archived rounds every abandoned call stayed
 * abandoned, and not one has ever answered late. So a call that has not
 * answered inside its budget is treated as lost rather than waited on.
 *
 * The budget is not the same for every call, and that is the point of passing
 * it in. The probe's third sheet timed out at thirty seconds on an insert whose
 * deck delta showed both slides had landed, so an insert gets sixty; a count is
 * two round trips and gets far less. A budget that fires on a call that worked
 * produces a false refusal, which on this host is the more expensive direction:
 * a refusal that never happened sends the reader hunting in the wrong file.
 */
export const BUDGET = {
  /** Reading or counting: cheap, and a slow one is a sick host. */
  read: 15000,
  /** Inserting a deck. Sixty because thirty was measured too short once. */
  insert: 60000,
  /** Removing slides this run added. */
  undo: 30000,
  /** Reading the whole package out of the host, which is slice by slice. */
  file: 90000,
} as const;

export class Timeout extends Error {
  constructor(
    readonly what: string,
    readonly ms: number,
  ) {
    super(`gave up waiting for: ${what} (after ${ms}ms)`);
    this.name = "Timeout";
  }
}

/**
 * Race a promise against a budget, and say so.
 *
 * The timer is cleared either way. Left running it keeps a Node process alive
 * after the work is done, which turns a passing test suite into one that hangs.
 *
 * **Every host call in the add-in comes through here**, and until now `what`
 * reached exactly one place: the message of a `Timeout` nobody sees unless the
 * call failed. So no successful call in this codebase named itself or its
 * duration, and there was no baseline against which a 41-second insert is
 * normal or alarming — every number in `BUDGET` was a guess with no measurement
 * behind it.
 *
 * Two rules the sibling project paid for, both obeyed here:
 *
 * - **The line is written BEFORE the call and named for what it knows.** Its
 *   own per-batch line was called `batch committed` and emitted one statement
 *   before the sync, so every stall left behind a line claiming the batch that
 *   killed it had committed; two separate analyses of that archive drew
 *   opposite wrong conclusions from it. `issued` is what this knows. The
 *   ordering is right and stays — a call that never answers has to be on the
 *   record while you are still waiting for it.
 * - **Both populations.** A field recorded only on failures cannot be compared
 *   against anything and is not yet a measurement. Four of that project's
 *   diagnostic fields died of exactly this, one round each.
 */
export function withTimeout<T>(work: Promise<T>, ms: number, what: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const bell = new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Timeout(what, ms)), ms);
  });
  const at = Date.now();
  trace("host", "issued", { call: what, budget: ms });
  return Promise.race([work, bell]).then(
    (value) => {
      if (timer !== undefined) clearTimeout(timer);
      trace("host", "answered", { call: what, ms: Date.now() - at });
      return value;
    },
    (err: unknown) => {
      if (timer !== undefined) clearTimeout(timer);
      // Three states, not two. A call that ran out of budget and a call that
      // raised are different facts about the host and want different next
      // steps, and collapsing them is how "the host got in the way" became the
      // sibling's least useful verdict.
      trace("host", err instanceof Timeout ? "gave up waiting" : "raised", {
        call: what,
        ms: Date.now() - at,
        ...(err instanceof Timeout ? {} : { error: err instanceof Error ? err.message : String(err) }),
      });
      throw err;
    },
  );
}
