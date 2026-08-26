/**
 * Every host call is bounded.
 *
 * On PowerPoint for the web a call that stops answering does not come back:
 * a sibling project recorded seventeen abandoned calls across nine rounds and
 * not one late answer. So a call that has not answered inside its budget is
 * treated as lost rather than waited on.
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
 * Race a promise against a budget.
 *
 * The timer is cleared either way. Left running it keeps a Node process alive
 * after the work is done, which turns a passing test suite into one that hangs.
 */
export function withTimeout<T>(work: Promise<T>, ms: number, what: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const bell = new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Timeout(what, ms)), ms);
  });
  return Promise.race([work, bell]).finally(() => {
    if (timer !== undefined) clearTimeout(timer);
  });
}
