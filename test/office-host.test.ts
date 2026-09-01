import { afterEach, describe, expect, it, vi } from "vitest";
import { ERROR_CHARS } from "../src/host/errors.js";
import { installFakeHost, uninstallFakeHost } from "./fixtures/host.js";

/**
 * The one call the pane makes before it has decided anything.
 *
 * `src/office` mostly cannot run in the suite — it calls Office.js — but
 * `hostSupports` is a single expression over `Office.context`, and what it does
 * when that object is not what it expects decides whether an unsupported host
 * gets a sentence or a blank pane.
 *
 * `ready()` uses it to render "this PowerPoint is too old", and
 * `hostEnvironment()` uses it to say which build was refused. Both run before
 * anything else, so a raise from either leaves the user with nothing to read on
 * exactly the host that needed the message — and nothing to report, because the
 * pane never drew.
 */
async function supportsWith(context: unknown): Promise<(v: string) => boolean> {
  vi.resetModules();
  (globalThis as unknown as { Office: unknown }).Office = { context };
  const mod = await import("../src/office/powerpoint.js");
  return mod.hostSupports;
}

afterEach(() => {
  delete (globalThis as unknown as { Office?: unknown }).Office;
});

describe("asking the host what it supports", () => {
  it("answers what the host says when the host can answer", () => {
    // The ordinary path, so the guard below is not just "always false".
    const asked: string[] = [];
    return supportsWith({
      requirements: {
        isSetSupported: (name: string, v: string) => {
          asked.push(`${name} ${v}`);
          return v === "1.2";
        },
      },
    }).then((supports) => {
      expect(supports("1.2")).toBe(true);
      expect(supports("1.10")).toBe(false);
      expect(asked).toEqual(["PowerPointApi 1.2", "PowerPointApi 1.10"]);
    });
  });

  it("answers false rather than throwing when there is no requirements object", async () => {
    const supports = await supportsWith({});
    expect(supports("1.2")).toBe(false);
  });

  it("answers false rather than throwing when there is no context at all", async () => {
    const supports = await supportsWith(undefined);
    expect(supports("1.2")).toBe(false);
  });

  it("answers false rather than throwing when the host raises", async () => {
    // A host that has the method and refuses the question. Same treatment: it
    // cannot answer, so it does not support it.
    const supports = await supportsWith({
      requirements: {
        isSetSupported: () => {
          throw new Error("no");
        },
      },
    });
    expect(supports("1.2")).toBe(false);
  });
});

/**
 * The two calls that ACT on the user's deck, driven against a fake host.
 *
 * `test/office-merge.test.ts` mocks both of these wholesale, so what they do
 * with a host's answer was checked by nothing. Both defects below were found
 * that way and neither is visible to a mock of the function holding it.
 */
describe("what an insert and an undo say when the host misbehaves", () => {
  async function host(options: Parameters<typeof installFakeHost>[0] = {}) {
    vi.resetModules();
    const fake = installFakeHost(options);
    const module = await import("../src/office/powerpoint.js");
    // From the SAME fresh registry. `resetModules` gives the host module its
    // own copy of every import, so a trace read through this file's top-level
    // import would be reading a different module's empty log.
    const trace = await import("../src/core/trace.js");
    return { fake, module, trace };
  }

  afterEach(() => {
    uninstallFakeHost();
  });

  it("names which HOST it is in, and never the word undefined", async () => {
    /**
     * Two defects in one line, both of which defeat the thing it exists for.
     * `host` was filled from `diagnostics.version`, so the field answering
     * "which application am I in" carried a build number; and the platform was
     * read through `String`, which turns an absent value into the STRING
     * "undefined" — truthy, so `environmentLine`'s `?? "unknown"` never fired
     * and the line read `platform: "undefined"`, which is the exact outcome
     * that fallback's own comment says it prevents.
     */
    const { module } = await host();
    const line = module.hostEnvironment();
    expect(line.host).toBe("PowerPoint");
    expect(line.officeVersion).toBe("16.0.fake");
    expect(line.platform).toBe("OfficeOnline");
  });

  it("says unknown when the host will not say, rather than the word undefined", async () => {
    vi.resetModules();
    (globalThis as unknown as { Office: unknown }).Office = {
      context: { requirements: { isSetSupported: () => true } },
    };
    const module = await import("../src/office/powerpoint.js");
    const line = module.hostEnvironment();
    expect(line.platform).toBe("unknown");
    expect(line.host).toBe("unknown");
    expect(line.officeVersion).toBe("unknown");
  });

  it("bounds what the host said, so a failed insert cannot put the deck on screen", async () => {
    // Office echoes an argument back through `debugInfo`, and the argument to
    // `insertSlidesFromBase64` is the ENTIRE merged deck as base64. Uncapped,
    // this sentence carried megabytes of the user's own merged rows to the
    // pane — and into a bug report, if they pasted it. `errors.ts` exists for
    // this and its callers were the paths that REJECT; this one is caught and
    // returned, so it never asked.
    const { module } = await host({
      onInsert: () => {
        throw new Error(`InvalidArgument: ${"A".repeat(2_000_000)}`);
      },
    });
    const outcome = await module.insertDeck("PKG", 2);
    expect(outcome.detail.length).toBeLessThanOrEqual(ERROR_CHARS + 64);
    expect(outcome.detail, "the reason itself is gone").toContain("InvalidArgument");
    expect(outcome.detail, "and what was dropped is counted rather than elided").toMatch(/more characters/);
  });

  it("does not spend the selection's budget walking the deck", async () => {
    /**
     * `deckSlideIds` pages the deck, and every page carries its own
     * `BUDGET.read`. It was nested inside one more of the same size, so the
     * outer budget bounded its own sub-budgets and fired whenever the TOTAL
     * crossed it — however promptly the host answered each call.
     *
     * A 600-slide deck at half a second per round trip is 28 calls and about
     * fifteen seconds of walking, so "use the slides I've selected" refused
     * every time with "gave up waiting". A 200-row merge over a three-slide
     * block leaves exactly that deck.
     *
     * The sync cost here is a tenth of the real one, so the test is quick; what
     * makes it bite is the ARITHMETIC — many bounded calls summing past one
     * bound — not the absolute numbers.
     */
    const slides = Array.from({ length: 600 }, (_, i) => `s${i}`);
    const { module, trace } = await host({ slides, syncMs: 8, selected: ["s3", "s4", "s5"] });
    trace.beginRun();
    const started = Date.now();
    const picked = await module.selectedBlock();
    const whole = Date.now() - started;
    expect(picked.ok, `refused: ${"why" in picked ? picked.why : ""}`).toBe(true);
    expect(picked).toMatchObject({ from: 4, to: 6 });

    // What the budget was actually asked to cover. A stopwatch could not tell
    // these apart — the WALK is the same work either way — so the assertion is
    // on how much of it the budgeted call was charged for. Nested, it is
    // charged for all of it; correct, it is charged for one batch.
    const charged = trace
      .traceLog()
      .entries.filter((e) => e.message === "answered" && e.data?.call === "reading the selected slides")
      .map((e) => Number(e.data?.ms ?? 0));
    expect(charged, "the call was never traced").toHaveLength(1);
    expect(charged[0], `charged ${charged[0]}ms of a ${whole}ms walk`).toBeLessThan(whole / 3);
  });

  it("names the slides the way the rail numbers them, never a 0-based index", async () => {
    // The pane speaks the thumbnail rail's numbering and nothing else. This
    // sentence reaches the user verbatim — "Nothing was removed — ${detail}" —
    // and said "from index 3" for slides the rail calls 4 to 9: a number one
    // before the ones actually touched, in a sentence about slides being
    // deleted. The refusal branch a few lines above it already got this right.
    const { fake, module } = await host({
      slides: ["u1", "u2", "u3", "m1", "m2", "m3", "m4", "m5", "m6"],
    });
    const outcome = await module.undoInsert(3, 6, "run-1");
    expect(outcome.removed).toBe(6);
    expect(outcome.detail, "an internal index reached the user").not.toContain("index");
    expect(outcome.detail).toContain("slides 4 to 9");
    expect(fake.slides).toEqual(["u1", "u2", "u3"]);
  });

  it("takes nothing on a REPEAT press when the slides cannot prove they are ours", async () => {
    /**
     * The same deck and the same call, with `requireProof`. This fake answers
     * the tag read with a null object for every slide, which is what a host
     * that has stopped answering looks like — and without `requireProof` the
     * fall-through takes the whole window, as the test above shows.
     *
     * That is the answer this add-in gave before tags existed and is right for
     * a FIRST press. On a later one it deletes: a press having happened is
     * itself proof the deck has changed shape, so the window a size clamp
     * produces can hold a slide the user made in between.
     */
    const { fake, module } = await host({
      slides: ["u1", "u2", "u3", "m1", "m2", "m3", "m4", "m5", "m6"],
    });
    const outcome = await module.undoInsert(3, 6, "run-1", { requireProof: true });
    expect(outcome.removed, "nothing may go without proof").toBe(0);
    expect(fake.slides, "and the deck is untouched").toHaveLength(9);
    expect(outcome.detail).toContain("could be shown to be this merge's");
    // The sentence may not claim WHICH of the two it was: a slide with no tag
    // and a slide the host would not answer for are the same `undefined` here.
    expect(outcome.detail).not.toContain("PowerPoint would not");
    expect(outcome.detail).not.toContain("carries this merge's mark");
    // And the pane may NOT treat this as terminal. This host has the API and
    // did not answer with it; the next press may well be answered, so the way
    // back stays on screen.
    expect(outcome.unprovable, "a host that can answer and did not").not.toBe(true);
  });

  it("takes nothing on a REPEAT press on a host with no tags either, and spares the slide made since", async () => {
    /**
     * `requireProof` was gated on `hostSupports("1.3")` for one commit, on the
     * reasoning that a 1.2 host can never answer a tag read, so demanding
     * proof of it leaves a card standing over slides no press can take.
     *
     * The gate is a slide deletion, and a guaranteed one rather than the
     * intermittent case the proof was added for: on EVERY 1.2 host the second
     * press falls through to the whole positional window, and that window is
     * exactly the one that can hold a slide the user made between the presses.
     *
     * The deck below is that sequence. Twelve slides of the user's, a
     * six-slide merge, a first press that took three; the user then deleted
     * one merged slide and made one of their own. The deck's growth still
     * equals what is owed, so every size clamp passes and `MINE` is the last
     * slide in the window.
     */
    const { fake, module } = await host({
      slides: [...Array.from({ length: 12 }, (_, i) => `u${i + 1}`), "m4", "m5", "MINE"],
      sets: ["1.1", "1.2"],
    });
    const outcome = await module.undoInsert(12, 3, "run-1", { requireProof: true });
    expect(outcome.removed, "a host that cannot prove anything may not delete on position alone").toBe(0);
    expect(outcome.disowned).toBe(3);
    expect(fake.slides, "the slide the user made between the presses").toContain("MINE");
    expect(fake.slides).toHaveLength(15);
    // This one IS terminal, and says so: there are no slide tags on this host
    // to prove anything with, so the same press answers the same way for ever.
    // It is the only shape the pane may withdraw the card on.
    expect(outcome.unprovable, "a host with no tags at all").toBe(true);
  });
});
