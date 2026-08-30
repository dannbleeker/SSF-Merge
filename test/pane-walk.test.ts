// @vitest-environment jsdom
/**
 * The pane under a random walk, through the REAL `main.ts`.
 *
 * `pane-wiring.test.ts` drives the sequences somebody thought of. This drives
 * the ones nobody did: press whatever is on the screen, in whatever order, for
 * a few hundred steps, and check after every single press that the screen is
 * still one a user could act on.
 *
 * Every defect this file's neighbour was written for is a SEQUENCE — a Back and
 * a Continue during a two-minute merge, a raise that left the button reading
 * "Merging…" for the rest of the session, a caret that jumped on every
 * keystroke. Those are found by walking, and the walk is cheap.
 *
 * **Seeded, and the seed is in the failure message.** A fuzz whose failure
 * cannot be reproduced is a fuzz nobody can act on: every walk here replays
 * exactly from its seed, and the assertion prints the seed and the step number
 * it died on.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const office = vi.hoisted(() => ({
  slideCount: vi.fn<() => Promise<number>>(),
  ready: vi.fn(() => ({ ok: true, detail: "fine" })),
  selectedBlock: vi.fn<() => Promise<unknown>>(),
  canReadSelection: vi.fn(() => true),
  hostEnvironment: vi.fn(() => ({
    build: "test",
    platform: "PC",
    host: "16.0.0",
    sets: ["1.2"],
    floor: "1.2",
    clearsFloor: true,
    deckSource: "file" as const,
    canSelect: true,
  })),
  insertTextAtCursor: vi.fn<(t: string) => Promise<unknown>>(),
  // Which deck is open. The undo crumb is kept in `localStorage`, which belongs
  // to the add-in's ORIGIN and is shared by every deck opened against it, so the
  // pane asks who is open before it answers for a crumb. Missing here, the walk
  // still reported all of its tests passing while nine unhandled errors came out
  // of the mocker and took the run's exit code with them.
  documentKey: vi.fn(() => "https://example-my.sharepoint.com/personal/x/Documents/deck.pptx"),
  inspectBlock: vi.fn<(r: { from: number; to: number }) => Promise<unknown>>(),
  runMerge: vi.fn<(r: unknown) => Promise<unknown>>(),
  undoMerge: vi.fn<(o: unknown) => Promise<unknown>>(),
}));
vi.mock("../src/office/powerpoint.js", () => office);
vi.mock("../src/office/merge.js", () => office);

const REPORT = {
  ok: true,
  fields: ["First", "Last"],
  imageFields: [],
  slideFields: [["First"], ["Last"], []],
  detail: "",
};
const OUTCOME = {
  ok: true,
  detail: "6 slides added after slide 12.",
  added: 6,
  deckAtStart: 12,
  runId: "r1",
  fields: ["First", "Last"],
  imageFields: [],
  slideFields: [["First"], ["Last"], []],
  unknownConditions: [],
};

let onReady: () => void = () => undefined;

async function openPane(): Promise<void> {
  document.body.innerHTML = '<header><b>SSF</b><span>Merge</span></header><div id="pane"></div>';
  vi.resetModules();
  (globalThis as unknown as { Office: unknown }).Office = {
    onReady: (cb: () => void) => {
      onReady = cb;
      return Promise.resolve();
    },
    context: {},
  };
  await import("../src/pane/main.js");
  onReady();
  await Promise.resolve();
}

function pane(): HTMLElement {
  return document.getElementById("pane") as HTMLElement;
}
async function settle(): Promise<void> {
  for (let i = 0; i < 12; i++) await Promise.resolve();
}

/**
 * A tiny seeded generator, so a failing walk is a walk somebody can re-run.
 *
 * `Math.random` would make this file report a defect nobody can reproduce,
 * which is worse than not looking: the next person re-runs it, sees green, and
 * concludes the report was noise.
 */
function rng(seed: number): () => number {
  let s = seed >>> 0;
  return () => {
    s = (Math.imul(s, 1664525) + 1013904223) >>> 0;
    return s / 0x100000000;
  };
}

/** Every label the pane's own `primary()` gives a call that is in flight. */
const RUNNING_LABELS = ["Reading the slides…", "Merging…", "Previewing…", "Removing…"];

/**
 * What has to be true of the screen after every single press.
 *
 * Deliberately about the SHAPE of the pane rather than about any one feature:
 * a walk cannot know what the right count is, but it can know that there is one
 * button, that the pane is not frozen, and that the screen explains itself.
 */
let crossChecks = 0;

function faults(): string[] {
  const out: string[] = [];
  const main = pane().querySelector("main");
  const primaries = Array.from(pane().querySelectorAll("button.primary"));

  // One filled button per screen, always last. The pane's own rule.
  if (primaries.length !== 1) out.push(`${primaries.length} primary buttons`);
  else if (main?.lastElementChild !== primaries[0]) out.push("the primary is not the last thing on the screen");

  const button = primaries[0] as HTMLButtonElement | undefined;
  if (button) {
    // Nothing is in flight once everything has settled, so no screen may still
    // be showing a call's label. This is the frozen pane, which is the defect
    // its neighbour file was written for.
    if (RUNNING_LABELS.includes(button.textContent ?? "")) out.push(`still says "${button.textContent ?? ""}"`);
    // NOT checked: "a disabled button is explained". Tried, and it is a
    // judgement rather than a shape — step 1's disabled button reads "Choose
    // the slides that repeat", which IS the explanation, and no rule over the
    // DOM tells that apart from a dead end. A walk that reports a false
    // positive on step 0 of every seed is a walk nobody runs twice.
  }

  // The button and the forecast are two renderings of ONE number, and this
  // repo has shipped them disagreeing before — a card reading "4 slides added
  // after slide 20" beside a button offering a different count. Both are on
  // screen together on the merge step, so a walk can simply read them.
  const offer = /^Add (\d+) slides?$/.exec(button?.textContent ?? "");
  const forecast = /^(\d+) slides? added after slide (\d+), leaving (\d+) slides? in the deck\.$/.exec(
    pane().querySelector(".card:not(.undo) p")?.textContent ?? "",
  );
  if (offer && forecast) {
    crossChecks++;
    if (offer[1] !== forecast[1]) out.push(`the button offers ${offer[1]} and the forecast says ${forecast[1]}`);
    // And the deck arithmetic in the same sentence has to add up.
    if (Number(forecast[2]) + Number(forecast[1]) !== Number(forecast[3])) {
      out.push(`${forecast[2]} + ${forecast[1]} is not ${forecast[3]}`);
    }
  }

  // One orange thing. The tick yields to a card that has taken the budget.
  if (pane().querySelector(".tick") && pane().querySelector(".card.undo")) out.push("two orange elements");

  // A step is always drawn.
  if (!pane().querySelector(".step-of")) out.push("no step line");

  // EVERY step that is not the last shows a way onward.
  //
  // The one defect a per-frame shape check nearly could not see. Pressing
  // "Preview the first row" left step 4 with "Back to fields" and "Remove the
  // preview" and nothing else — the word "merge" appeared nowhere on the screen
  // — and the route on was to work out that clearing the preview was the route.
  // Every rule above passed on it: there was one primary, it was last, it was
  // enabled, and it said something true.
  //
  // Deliberately NOT gated on the control being enabled, which is the trap the
  // note above about disabled buttons describes. Step 1 opens with "Choose the
  // slides that repeat", disabled, and that IS the way to step 2 — the door is
  // there and is not open yet. So this asks whether the door EXISTS:
  // `data-advances` on the primary, or a `data-forward` link beside it.
  const stepLine = /^Step (\d+) of (\d+)/.exec(pane().querySelector(".step-of")?.textContent ?? "");
  if (stepLine && stepLine[1] !== stepLine[2]) {
    const onward = pane().querySelector("[data-advances], [data-forward]");
    if (!onward) out.push(`step ${stepLine[1]} of ${stepLine[2]} offers no way onward`);
  }
  return out;
}

/** Everything on the screen a user could press or type into, right now. */
function moves(): (() => void)[] {
  const out: (() => void)[] = [];
  for (const node of Array.from(pane().querySelectorAll("button, [data-back], [data-forward]"))) {
    if (node instanceof HTMLButtonElement && node.disabled) continue;
    out.push(() => (node as HTMLElement).click());
  }
  for (const node of Array.from(pane().querySelectorAll("input[type=checkbox]"))) {
    out.push(() => (node as HTMLElement).click());
  }
  for (const node of Array.from(pane().querySelectorAll("input[type=text], textarea"))) {
    const el = node as HTMLInputElement;
    for (const value of ["4", "6", "First\tLast\nAda\tLovelace\nGrace\tHopper", ""]) {
      out.push(() => {
        el.value = value;
        el.dispatchEvent(new Event("input", { bubbles: true }));
      });
    }
  }
  for (const node of Array.from(pane().querySelectorAll("select"))) {
    const el = node;
    for (const option of Array.from(el.options)) {
      out.push(() => {
        el.value = option.value;
        el.dispatchEvent(new Event("input", { bubbles: true }));
      });
    }
  }
  return out;
}

/**
 * Walk to the merge step the way a user does, before the random walk starts.
 *
 * Without this the walks never get there: reaching step 5 takes a specific
 * six-press sequence, and a uniform choice over the thirty-odd controls on
 * screen finds it about never. The first version of this file measured that
 * the hard way — a cross-check between the button's count and the forecast
 * card fired ZERO times across five walks, which is the vacuous measurement
 * this repo keeps catching itself at.
 */
async function reachMerge(): Promise<void> {
  await openPane();
  await settle();
  const type = (name: string, value: string): void => {
    const node = pane().querySelector(`[data-field="${name}"]`) as HTMLInputElement;
    node.value = value;
    node.dispatchEvent(new Event("input", { bubbles: true }));
  };
  const primary = (): HTMLButtonElement => pane().querySelector("button.primary") as HTMLButtonElement;
  type("from", "4");
  type("to", "6");
  primary().click();
  await settle();
  type("paste", "First\tLast\nAda\tLovelace\nGrace\tHopper");
  primary().click();
  await settle();
  primary().click();
  await settle();
  (pane().querySelector("[data-forward]") as HTMLElement).click();
  await settle();
}

beforeEach(() => {
  localStorage.clear();
  office.slideCount.mockReset().mockResolvedValue(12);
  office.canReadSelection.mockReset().mockReturnValue(true);
  office.insertTextAtCursor.mockReset().mockResolvedValue({ ok: true });
  office.selectedBlock.mockReset().mockResolvedValue({ ok: true, from: 4, to: 6 });
  // Always answering, because a walk that leaves a call outstanding is
  // measuring the mock rather than the pane. The refusals get their own walk.
  office.inspectBlock.mockReset().mockResolvedValue(REPORT);
  office.runMerge.mockReset().mockResolvedValue(OUTCOME);
  office.undoMerge.mockReset().mockResolvedValue({ removed: 6, detail: "6 removed" });
});
afterEach(() => {
  vi.unstubAllGlobals();
});

describe("the pane survives a random walk", () => {
  // Fixed seeds rather than a count: the same five walks run on every machine
  // and in every CI job, so a red one is a red one everywhere.
  for (const seed of [1, 7, 42, 1234, 99991]) {
    it(`holds its shape on seed ${seed}`, async () => {
      const next = rng(seed);
      await openPane();
      await settle();
      for (let step = 0; step < 120; step++) {
        const options = moves();
        if (options.length === 0) break;
        options[Math.floor(next() * options.length)]?.();
        await settle();
        const found = faults();
        expect(found, `seed ${seed}, step ${step}: ${found.join("; ")}`).toEqual([]);
      }
    });
  }

  it("holds its shape when every host call refuses", async () => {
    // The other half: every host call fails. A refusal is what leaves a pane
    // frozen if a `finally` is missing anywhere.
    //
    // Which calls RAISE and which ANSWER is not a free choice here — it is the
    // office layer's own contract, written at each function. `selectedBlock`
    // and `insertTextAtCursor` promise never to reject, in comments that name
    // this exact defect ("the pane awaits this from a click handler, and a
    // rejection there is an unhandled one"), and the pane has no catch around
    // either. Mocking them as rejecting measures the mock: it produces an
    // unhandled rejection that says nothing about the pane. So they answer
    // `{ ok: false }`, which is what they really do.
    office.inspectBlock.mockReset().mockRejectedValue(new Error("the host said no"));
    office.runMerge.mockReset().mockRejectedValue(new Error("the host said no"));
    office.undoMerge.mockReset().mockRejectedValue(new Error("the host said no"));
    office.selectedBlock.mockReset().mockResolvedValue({ ok: false, why: "the host said no" });
    office.insertTextAtCursor.mockReset().mockResolvedValue({ ok: false, why: "the host said no" });
    const next = rng(2026);
    await openPane();
    await settle();
    for (let step = 0; step < 120; step++) {
      const options = moves();
      if (options.length === 0) break;
      options[Math.floor(next() * options.length)]?.();
      await settle();
      const found = faults();
      expect(found, `refusing walk, step ${step}: ${found.join("; ")}`).toEqual([]);
    }
  });

  // Starting FROM the merge step, because that is where the interesting
  // controls are — the row picker, the conditions, the blank-cell answer, the
  // undo card — and where the two numbers this walk cross-checks are both on
  // screen. A walk from a fresh pane reaches it about never.
  for (const seed of [3, 11, 777]) {
    it(`holds its shape walking from the merge step on seed ${seed}`, async () => {
      const next = rng(seed);
      await reachMerge();
      for (let step = 0; step < 120; step++) {
        const options = moves();
        if (options.length === 0) break;
        options[Math.floor(next() * options.length)]?.();
        await settle();
        const found = faults();
        expect(found, `seed ${seed}, step ${step}: ${found.join("; ")}`).toEqual([]);
      }
    });
  }

  it("actually compared the button against the forecast", () => {
    // The vacuity guard. Every assertion above passes on a walk that never
    // reached the screen carrying both numbers, and the first version of this
    // file did exactly that — zero comparisons across five walks, all green.
    expect(crossChecks).toBeGreaterThan(20);
  });
});
