/**
 * An ordered record of what a run actually did, for the runs nobody watched.
 *
 * A merge is one click and then up to two and a half minutes of silence:
 * `BUDGET.file` allows ninety seconds to read the template and `BUDGET.insert`
 * sixty to hand the package over. If it goes wrong the user has a sentence and
 * we have nothing — no idea which call was in flight, how long the ones before
 * it took, or whether the host was slow throughout or fine until the last step.
 *
 * `console.log` does not answer that. A task pane on the web has no devtools
 * the user can open, and a host that takes the tab down takes the console with
 * it. So the record is a value the pane can hand back.
 *
 * **Deliberately small.** The sibling project's equivalent is a 2000-entry ring
 * with slice arithmetic, two histograms and a per-operation mark, and it needs
 * all of that: a real run there emits 276 entries and pretty-prints to 160 KB.
 * A merge emits about ten. A capped array cleared at the start of each run is
 * the whole mechanism, and building the ring first would be an instrument
 * shaped by somebody else's failures.
 *
 * Lives in `src/core` because the coverage config's `include` is a fixed list
 * of three globs — a new top-level directory would be measured by nothing, and
 * an uncounted module is how a threshold quietly stops meaning anything.
 */

/** One thing that happened, and when. */
export interface TraceEntry {
  /**
   * Milliseconds since `beginRun`, so the log reads as a timeline.
   *
   * Relative to ONE origin per run, and everything that stamps a time into the
   * same artefact must use it. The sibling stamped its probe samples from a
   * different start and every cross-reference in the file was off by a constant
   * — samples appearing to arrive before the pass that produced them.
   */
  ms: number;
  /** Which part spoke: "merge", "host", "pane". */
  scope: string;
  /** What happened, named for what was KNOWN when the line was written. */
  message: string;
  /** Structured extras. A breadcrumb, not a heap dump. */
  data?: Record<string, unknown>;
}

/**
 * Entries kept before the oldest is dropped.
 *
 * Generous for a run that emits ten, and a bound rather than a hope: a merge
 * over four hundred rows that somehow traced per row must not grow an unbounded
 * array inside a task pane that is already struggling.
 */
export const MAX_ENTRIES = 500;

/** How much of one payload value may appear on a line. */
const VALUE_CHARS = 200;

let entries: TraceEntry[] = [];
let dropped = 0;
let startedAt = 0;
let running = false;
let watcher: ((e: TraceEntry) => void) | undefined;

/**
 * Start recording, discarding whatever came before.
 *
 * Per RUN, not per session. A log that spans two merges pairs one run's numbers
 * with another run's failures, which is the wrong turn that costs an hour.
 */
export function beginRun(): void {
  entries = [];
  dropped = 0;
  startedAt = Date.now();
  running = true;
}

/** Whether a run is being recorded. */
export function tracing(): boolean {
  return running;
}

/**
 * Milliseconds since `beginRun`, or null when nothing is being recorded.
 *
 * Exported so anything else stamping a time into the same run log shares this
 * origin rather than starting its own.
 */
export function elapsed(): number | null {
  return running ? Date.now() - startedAt : null;
}

/**
 * Record one thing.
 *
 * The payload is COPIED. A log is a record of what was true at the moment it
 * was written, and holding the caller's object lets a later mutation rewrite
 * history — in a file somebody may already be reading as fact.
 */
export function trace(scope: string, message: string, data?: Record<string, unknown>): void {
  if (!running) return;
  const entry: TraceEntry = { ms: Date.now() - startedAt, scope, message, ...(data ? { data: { ...data } } : {}) };
  entries.push(entry);
  // A broken window must never cost the record the entry it was writing —
  // least of all during the failure it is there to photograph.
  try {
    watcher?.(entry);
  } catch {
    /* the window is broken; the record is not */
  }
  if (entries.length > MAX_ENTRIES) {
    entries.shift();
    dropped++;
  }
}

/**
 * Watch each entry as it is recorded. Pass undefined to stop.
 *
 * The log is the record; this is the window, and they are different needs. A
 * run log can only be handed over once the run ENDS, and the runs worth
 * explaining are the ones that never do. What is already on screen survives a
 * host that takes the pane down, and can be read or copied.
 *
 * One subscriber, replaced rather than accumulated: there is one pane.
 */
export function onTrace(cb: ((e: TraceEntry) => void) | undefined): void {
  watcher = cb;
}

/** One payload value as it appears on a line, or undefined to leave it off. */
export function formatValue(v: unknown): string | undefined {
  // `undefined` and "" both mean "nothing to say", and `key=` with nothing
  // after it is noise. `0`, `false` and `null` are kept: they are ANSWERS, and
  // `filled=0` is the entire finding of a merge that matched no placeholders.
  if (v === undefined || v === "") return undefined;
  if (v === null) return "null";
  // A function in a payload is a caller's slip, and `String(fn)` prints its
  // whole source into a breadcrumb. Named, not rendered.
  if (typeof v === "function") return "(function)";
  // Each primitive named rather than one `String(v)` over a union. A payload
  // arrives as `unknown` — the whole point is that a call site can hand this
  // anything — and `String` over that reaches Object's default
  // stringification, which renders "[object Object]": a line that occupies
  // space and answers nothing.
  // Capped like every other value. A string payload is the likeliest one to
  // be enormous — an Office error message carries a `debugInfo` blob that
  // echoes the argument back, and the argument to `insertSlidesFromBase64` is
  // the whole merged deck.
  if (typeof v === "string") return cap(v);
  if (typeof v === "number" || typeof v === "boolean" || typeof v === "bigint") return String(v);
  if (typeof v === "symbol") return v.toString();
  if (typeof v !== "object") return "(unprintable)";
  // A SHAPE rather than `String(v)` when JSON cannot render it: an object's
  // default stringification is "[object Object]", which occupies a line and
  // says nothing. A circular payload is a caller's mistake, not a reason to
  // lose the entry it was attached to.
  const shape = (): string => (Array.isArray(v) ? `[${v.length} items]` : "{…}");
  try {
    return cap(JSON.stringify(v) ?? shape());
  } catch {
    return shape();
  }
}

/** One value, cut to `VALUE_CHARS`. A breadcrumb should be short. */
function cap(text: string): string {
  return text.length > VALUE_CHARS ? `${text.slice(0, VALUE_CHARS - 1)}…` : text;
}

/**
 * One entry as one line.
 *
 * The ONE rendering, shared by whatever shows the log and whatever saves it.
 * The screen and the file must never describe the same run differently, and
 * they will the moment there are two formatters.
 */
export function formatTraceLine(e: TraceEntry): string {
  const bits = Object.entries(e.data ?? {})
    .map(([k, v]) => {
      const text = formatValue(v);
      return text === undefined ? undefined : `${k}=${text}`;
    })
    .filter((b): b is string => b !== undefined)
    .join(" ");
  const secs = String(Math.round(e.ms / 100) / 10).padStart(6);
  return `${secs}s  ${e.scope}  ${e.message}${bits ? `  ${bits}` : ""}`;
}

/** The run so far, oldest first, plus how many entries fell off the front. */
export function traceLog(): { entries: TraceEntry[]; dropped: number } {
  return {
    entries: entries.map((e) => ({ ...e, ...(e.data ? { data: { ...e.data } } : {}) })),
    dropped,
  };
}

/** The run so far as text, which is what a person copies out of the pane. */
export function traceText(): string {
  const { entries: list, dropped: lost } = traceLog();
  const head = lost > 0 ? [`(${lost} earlier line(s) dropped)`] : [];
  return [...head, ...list.map(formatTraceLine)].join("\n");
}
