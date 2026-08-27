/**
 * What an observation means, decided away from the host.
 *
 * A probe that reasons inside a `PowerPoint.run` callback is a probe whose
 * conclusions cannot be tested. Every judgement this one makes lives here, as a
 * pure function over numbers and strings, so the rules can go red in CI while
 * the round trip through a real PowerPoint stays a matter of reading values
 * out.
 *
 * The wording matters as much as the logic. A sibling project spent two
 * sessions reasoning about answers that could equally have been about the host
 * or about the thing asking, so every verdict here either names a mechanism or
 * says plainly that it cannot tell.
 */

export type Verdict = "yes" | "no" | "unknown" | "threw";

export interface InsertObservation {
  /** Deck size before the call, measured. */
  before: number;
  /** Deck size after the call, measured in its own context. */
  after: number;
  /** How many slides the file being inserted actually contains. */
  expected: number;
  /** The error, if the call threw. */
  error?: string;
}

export interface InsertVerdict {
  verdict: Verdict;
  landed: number;
  detail: string;
}

/**
 * Whether an insert did what it said.
 *
 * The delta is the evidence, never the absence of an error: this host family
 * has accepted slide adds, deck inserts and tag writes and performed none of
 * them. A partial landing is called out separately from nothing landing,
 * because they point at different things.
 */
export function insertVerdict(o: InsertObservation): InsertVerdict {
  const landed = o.after - o.before;
  if (o.error !== undefined) {
    // A call can raise and still have done the work. The third real sheet timed
    // out on an insert whose deck delta was exactly the two slides it was
    // asked for, and reading the error as decisive turned one late answer into
    // three false statements: that our package was refused, that the collision
    // arm disagreed with the fresh one, and that the theme was the difference.
    // The docstring above already said the delta is the evidence; this is the
    // code finally agreeing with it.
    if (landed === o.expected) {
      return {
        verdict: "yes",
        landed,
        detail: `all ${o.expected} slide(s) landed, but not before the probe stopped waiting (${o.error}). That is our budget being short, not the host refusing.`,
      };
    }
    return {
      verdict: "threw",
      landed,
      detail: `the call threw: ${o.error}${landed === 0 ? "" : `, and ${landed} slide(s) landed anyway`}`,
    };
  }
  if (landed === o.expected) {
    return { verdict: "yes", landed, detail: `all ${o.expected} slide(s) landed` };
  }
  if (landed === 0) {
    return {
      verdict: "no",
      landed,
      detail: `the call raised nothing and the deck did not grow, so ${o.expected} slide(s) were dropped silently`,
    };
  }
  return {
    verdict: "no",
    landed,
    detail: `${landed} of ${o.expected} slide(s) landed, which is a partial insert rather than a refusal`,
  };
}

/**
 * The headline for question one, from both arms.
 *
 * Asking only the fresh-id arm cannot separate "office-js#6105 does not
 * reproduce here" from "this host refuses every insert", so the collision arm
 * runs in the same session and the reading is taken from the pair. That is the
 * partner-question rule: a probe answer that admits two readings is not an
 * answer.
 */
export function creationIdReading(fresh: InsertVerdict, collision: InsertVerdict): string {
  if (fresh.verdict === "yes" && collision.verdict !== "yes") {
    return "CONFIRMED: a fresh creation id inserts and a duplicated one does not. The engine's rewrite is what makes cloning work, and office-js#6105 reproduces here.";
  }
  if (fresh.verdict === "yes" && collision.verdict === "yes") {
    return "Both arms landed. The creation id rewrite is harmless but this host does not need it: office-js#6105 does not reproduce here. Keep the rewrite, since the hosts that do need it are the ones we cannot test.";
  }
  if (fresh.verdict !== "yes" && collision.verdict !== "yes") {
    return "BLOCKING: neither arm landed, so this says nothing about creation ids and everything about insertion. Read the errors before building on the package path.";
  }
  return "Unexpected: the duplicated id landed and the fresh one did not. Do not build on either reading; re-run before drawing any conclusion.";
}

/** The value the probe writes into the fixture decks' package tags. */
export const PROBE_RUN_TAG = "probe-run";

export interface TagObservation {
  /** What the host said the tag holds, if it found one. */
  value?: string;
  error?: string;
  /**
   * How many slides the insert that was supposed to carry this tag landed.
   *
   * Load-bearing, and the reason this function exists. The probe reads the tag
   * off the LAST slide in the deck, which is the inserted one only if the
   * insert worked. When it did not, that read lands on a slide the user owns,
   * which has never carried our tag and never will.
   */
  insertLanded: number;
}

/**
 * Whether a tag written into the package survives into the host's object model.
 *
 * A never-asked question is not an answer, and reporting one as `no` is worse
 * than reporting nothing: the first sheet this probe took said "the metadata
 * scheme needs rethinking" on the strength of a read that had fallen on the
 * user's own title slide, because the insert in front of it had thrown. The
 * scheme was never tested. That is the same mistake a sibling project's
 * contract gate made with `no-scratch-slide`, and it is guarded here rather
 * than left to the reader to notice.
 */
export function tagVerdict(o: TagObservation): { verdict: Verdict; detail: string } {
  if (!(o.insertLanded > 0)) {
    return {
      verdict: "unknown",
      detail:
        "NOT ASKED — the slide that would carry the tag never landed, so this read fell on a slide the probe did not write. An absent tag here is evidence about nothing.",
    };
  }
  if (o.error !== undefined) return { verdict: "threw", detail: `the read threw: ${o.error}` };
  if (o.value === PROBE_RUN_TAG) {
    return { verdict: "yes", detail: "the whole metadata scheme works, and no tag write is needed in the host" };
  }
  if (o.value === undefined) {
    return {
      verdict: "no",
      detail: "the slide landed and the host did not find the tag, so the metadata scheme needs rethinking",
    };
  }
  return { verdict: "unknown", detail: `the tag is there holding ${JSON.stringify(o.value)}, which nothing wrote` };
}

/**
 * Whose fault a refused insert is.
 *
 * `insertSlidesFromBase64` answering `InvalidArgument` admits two readings that
 * matter in opposite directions: this repo generates a package PowerPoint will
 * not take, or this host takes no package at all. One is a morning's work and
 * the other ends the package path.
 *
 * The control arm is the presentation's OWN bytes, read back through
 * `getFileAsync`. That deck is a package PowerPoint wrote seconds earlier, so
 * it cannot be malformed, and a host that refuses it is refusing insertion
 * itself. This is the partner-question rule: do not reason about which of two
 * readings holds, ask a question only one of them survives.
 */
export function insertionBlame(ours: Verdict, self: Verdict): string {
  if (ours === "yes") return "The insert path works. Whose package it is does not arise.";
  if (self === "yes") {
    return "OURS: this host inserted a deck PowerPoint itself wrote and refused the one we generate. The package writer is at fault; nothing in the failing arms is a fact about the host.";
  }
  if (self === "unknown") {
    return "CANNOT TELL: the control arm never ran, so a refused insert is equally our package or this host. Re-run before concluding anything.";
  }
  return "THE HOST: it refused a deck it wrote itself, so nothing can be inserted here and the package path is blocked for reasons no change in this repo can reach.";
}

/**
 * The two experiments questions three and four run, and what each model
 * predicts. Here rather than in the snippet builder so the reader cannot come
 * to expect a string the probe stopped producing: a guard in
 * `test/probe.test.ts` asserts the generated snippet carries these.
 */
export const Q3 = {
  text: "Hello NAME here",
  /** The one targeted write replaces NAME with Ada. */
  want: "Hello Ada here",
} as const;

export const Q4 = {
  text: "AAA-BBB",
  /**
   * Five characters over AAA, then three over what WAS BBB, both queued in one
   * batch at offsets taken from the ORIGINAL string.
   *
   * If the host evaluates both against the text as it was, the second write
   * lands on BBB. If it evaluates the second against the first one's result,
   * the same offset is three characters to the left. Neither model runs off the
   * end, so a throw here is about something other than the offsets.
   */
  independent: "XXXXX-2",
  shifted: "XXXX2BB",
} as const;

export interface SubstringObservation {
  /** Text of the shape before the write. */
  before: string;
  /** Text after the write. */
  after: string;
  /** What the write was supposed to produce. */
  want: string;
  /** Whether the styled span was still styled afterwards. */
  boldAfter?: boolean;
  error?: string;
}

/**
 * Whether a substring write kept the formatting around it.
 *
 * Two things are asked at once on purpose, because a write that lands in the
 * wrong place and a write that flattens the run look the same in a screenshot.
 */
export function substringVerdict(o: SubstringObservation): { verdict: Verdict; detail: string } {
  if (o.error !== undefined) return { verdict: "threw", detail: `the write threw: ${o.error}` };
  if (o.after !== o.want) {
    return {
      verdict: "no",
      detail: `the text came out as ${JSON.stringify(o.after)} rather than ${JSON.stringify(o.want)}, so the offsets are not what the caller thinks`,
    };
  }
  if (o.boldAfter === undefined)
    return {
      verdict: "unknown",
      detail: "the text is right; the host would not say whether the run kept its formatting",
    };
  return o.boldAfter
    ? { verdict: "yes", detail: "the text is right and the styled run is still styled, so a targeted write is safe" }
    : {
        verdict: "no",
        detail: "the text is right but the styling is gone, so a targeted write flattens the run it lands in",
      };
}

/**
 * Whether two writes queued in one batch interfere.
 *
 * The answer decides whether replacements must be queued right to left or can
 * be issued in any order, which is the difference between a correct in-place
 * preview and one that silently scrambles a sentence.
 */
export function offsetVerdict(
  after: string,
  ifIndependent: string,
  ifShifted: string,
): { verdict: Verdict; detail: string } {
  if (after === ifIndependent) {
    return {
      verdict: "yes",
      detail: "both writes used the ORIGINAL offsets, so replacements may be queued in any order",
    };
  }
  if (after === ifShifted) {
    return {
      verdict: "no",
      detail: "the second write saw the first one's result, so replacements must be queued right to left",
    };
  }
  return {
    verdict: "unknown",
    detail: `neither model predicts ${JSON.stringify(after)}. Record it and do not guess: queue right to left, which is correct under both.`,
  };
}

export interface ExportPartsObservation {
  /** False when the host has no `exportAsBase64Presentation` at all. */
  supported?: boolean;
  sourceParts?: number;
  exportParts?: number;
  /** In the file-route package and not in the export. Capped by the probe. */
  missing?: string[];
  sourceHasAuthors?: boolean;
  exportHasAuthors?: boolean;
  sourceComments?: number;
  exportComments?: number;
  error?: string;
}

/**
 * Whether the export this add-in reads its template through keeps every part.
 *
 * office-js#6867 reports `Slide.exportAsBase64` omitting modern comments and
 * `ppt/authors.xml`. A sibling project triaged that as no exposure and was
 * right to — it calls the API for a PICTURE of a slide. Here the
 * presentation-level export is how the TEMPLATE IS READ before it is cloned, so
 * a part the export drops is a part every merged slide is missing, silently, in
 * a file that opens cleanly.
 *
 * THREE STATES, and the third is the one that matters. A deck with no comments
 * cannot answer this question, and an export with no `authors.xml` taken from
 * such a deck is not evidence that the export dropped anything. Reporting that
 * as "keeps everything" would be the `no-scratch-slide` mistake in a new place:
 * a question the run could not put, recorded as an answer.
 */
export function exportPartsVerdict(o: ExportPartsObservation): { verdict: Verdict; detail: string } {
  if (o.supported === false) {
    return {
      verdict: "unknown",
      detail:
        "NOT ASKED — this host has no exportAsBase64Presentation, so the template is read through getFileAsync and this question does not arise here.",
    };
  }
  if (o.error !== undefined) return { verdict: "threw", detail: `the comparison threw: ${o.error}` };
  if (o.sourceParts === undefined || o.exportParts === undefined) {
    return { verdict: "unknown", detail: "NOT ASKED — this sheet predates the arm." };
  }
  const hadAuthors = o.sourceHasAuthors === true;
  const hadComments = (o.sourceComments ?? 0) > 0;
  if (!hadAuthors && !hadComments) {
    return {
      verdict: "unknown",
      detail: `NOT ASKED — this deck carries no comments and no authors part, so there was nothing for the export to drop. Re-run on a deck with comments. (${o.sourceParts} parts in, ${o.exportParts} out.)`,
    };
  }
  const lostAuthors = hadAuthors && o.exportHasAuthors !== true;
  const lostComments = hadComments && (o.exportComments ?? 0) < (o.sourceComments ?? 0);
  if (lostAuthors || lostComments) {
    const lost = [
      lostAuthors ? "ppt/authors.xml" : "",
      lostComments ? `${(o.sourceComments ?? 0) - (o.exportComments ?? 0)} comment part(s)` : "",
    ]
      .filter(Boolean)
      .join(" and ");
    return {
      verdict: "yes",
      detail: `the export DROPS ${lost} — office-js#6867 reaches the presentation-level call too, so a merged deck loses them. ${o.sourceParts} parts in, ${o.exportParts} out.`,
    };
  }
  return {
    verdict: "no",
    detail: `the export kept the comments and the authors part this deck carries (${o.sourceParts} parts in, ${o.exportParts} out${(o.missing ?? []).length > 0 ? `, ${(o.missing ?? []).length} other part(s) not carried over` : ""}).`,
  };
}

export interface TornInsert {
  /** Rows whose every slide landed. */
  complete: number;
  /** Rows that got some slides but not all of them. */
  torn: number;
  /** Rows that got nothing at all. */
  absent: number;
  /** The first incomplete row, 0-based, or undefined when every row is whole. */
  firstIncomplete?: number;
  detail: string;
}

/**
 * A partial insert read in ROWS, which is the unit a mail merge has.
 *
 * `insertVerdict` grades slides, and for the probe that is right — it inserts a
 * two-slide fixture and the slides are the whole question. For a merge it is
 * the wrong unit and says almost nothing: **719 of 720** means one row's
 * three-slide block became two, every later row still looks correct, and the
 * user finds it at slide 141 with no idea it was ever going to be there. "One
 * of your 240 rows is incomplete" is a sentence somebody can act on.
 *
 * `slidesPerRecord` is taken from the plan, in plan order, one entry per row
 * that produced anything.
 *
 * **THE PREFIX ASSUMPTION IS NAMED RATHER THAN HIDDEN.** This walks the rows in
 * order and stops where the slides ran out, which is the reading if the host
 * truncated the package. Nothing establishes that a partial insert truncates
 * rather than dropping a slide from the middle — no round has produced one —
 * so the COUNT of whole rows is exact only under that reading.
 *
 * It is stated and not measured because it changes no advice: the answer to a
 * torn insert is to take the slides back and run it again, whichever row tore.
 * Where every row produces the same number of slides — the ordinary case, since
 * only a condition varies it — the count is position-independent anyway, and
 * the assumption buys nothing but the row's index.
 */
export function tornInsert(slidesPerRecord: number[], landed: number): TornInsert {
  const total = slidesPerRecord.reduce((a, b) => a + b, 0);
  const rows = slidesPerRecord.length;
  if (landed >= total) {
    return { complete: rows, torn: 0, absent: 0, detail: `all ${rows} row(s) landed complete` };
  }
  let seen = 0;
  let complete = 0;
  let firstIncomplete: number | undefined;
  for (const [i, slides] of slidesPerRecord.entries()) {
    if (seen + slides <= landed) {
      seen += slides;
      complete++;
      continue;
    }
    firstIncomplete = i;
    break;
  }
  // Whether the first incomplete row got SOME of its slides or none. A row with
  // nothing is missing; a row with two of three is torn, and torn is the worse
  // of the two because it looks finished.
  const got = landed - seen;
  const torn = firstIncomplete !== undefined && got > 0 ? 1 : 0;
  const absent = rows - complete - torn;
  const detail =
    torn > 0
      ? `${complete} of ${rows} row(s) landed complete; row ${(firstIncomplete ?? 0) + 1} got ${got} of its ${slidesPerRecord[firstIncomplete ?? 0] ?? 0} slide(s)` +
        (absent > 0 ? `, and ${absent} row(s) got nothing` : "")
      : `${complete} of ${rows} row(s) landed complete; ${absent} row(s) got nothing`;
  return { complete, torn, absent, ...(firstIncomplete !== undefined ? { firstIncomplete } : {}), detail };
}
