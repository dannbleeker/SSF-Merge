// @vitest-environment jsdom
/**
 * The pane driven the way a user drives it, through the REAL `main.ts`.
 *
 * Everything else in this suite checks the pane's decisions as pure functions,
 * which is the right shape for them and is blind to the one thing `main.ts`
 * owns: the order things happen in when a host call is out for ninety seconds
 * and the user keeps clicking. Every guard in this file is for a defect an
 * adversarial review reproduced against the committed code — a merge that
 * raised and left the button reading "Merging…" for the rest of the session, a
 * second merge started over the first through Back and Continue, a caret that
 * jumped to the end of the box on every keystroke.
 *
 * Only the two Office-touching modules are mocked. The state machine, the
 * renderer and the wiring are the real ones, because a copy of `merge()` in a
 * test file is a copy that stops matching.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const readTemplateRefusal = "PowerPoint would not name every slide between 4 and 6.";

const office = vi.hoisted(() => ({
  slideCount: vi.fn<() => Promise<number>>(),
  ready: vi.fn(() => ({ ok: true, detail: "fine" })),
  selectedBlock: vi.fn<() => Promise<unknown>>(),
  canReadSelection: vi.fn(() => true),
  // The environment line the run emits after its mark. Mocked flat rather than
  // omitted: this module is mocked EXPLICITLY, so a new export is a loud
  // failure rather than an undefined the pane trips over at run time.
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
  // Which deck is open. The undo crumb lives in `localStorage`, which belongs to
  // the add-in's ORIGIN and is shared by every deck opened against it, so the
  // crumb records the deck it was written on and is only answered for a match.
  // One stable value here means every wiring test below speaks about the same
  // document; the refusals themselves are pinned in `test/crumb.test.ts`.
  documentKey: vi.fn(() => "https://example-my.sharepoint.com/personal/x/Documents/deck.pptx"),
  inspectBlock: vi.fn<(r: { from: number; to: number }) => Promise<unknown>>(),
  runMerge: vi.fn<(r: unknown) => Promise<unknown>>(),
  undoMerge: vi.fn<(o: unknown) => Promise<unknown>>(),
}));

vi.mock("../src/office/powerpoint.js", () => ({
  ready: office.ready,
  slideCount: office.slideCount,
  selectedBlock: office.selectedBlock,
  canReadSelection: office.canReadSelection,
  hostEnvironment: office.hostEnvironment,
  insertTextAtCursor: office.insertTextAtCursor,
  documentKey: office.documentKey,
}));
vi.mock("../src/office/merge.js", () => ({
  inspectBlock: office.inspectBlock,
  runMerge: office.runMerge,
  undoMerge: office.undoMerge,
}));

/** A promise plus the handles to settle it, so a call can be held open. */
function deferred<T>() {
  let resolve!: (v: T) => void;
  let reject!: (e: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

const REPORT = { ok: true, detail: "2 placeholders in slides 4 to 6.", fields: ["First", "Last"] };
const OUTCOME = {
  ok: true,
  detail: "6 slides added after slide 12.",
  added: 6,
  // The ordinary case: the deck grew by what the package held, so the run can
  // account for every new slide and may offer to take them back.
  accountable: true,
  deckAtStart: 12,
  runId: "r1",
  fields: ["First", "Last"],
  unknownConditions: [],
};

/** How many steps the wizard has, for the walk-back helpers. */
const STEP_COUNT = 5;

let onReady: () => void;

/** Load the pane fresh, run its Office.onReady, and hand back the root. */
async function openPane(): Promise<HTMLElement> {
  // The HEADER as well as the pane, because `taskpane.html` has one and things
  // are drawn into it. Boot the harness with a bare `#pane` and `showBuild`
  // finds no header, returns, and the guard below passes on a page that is not
  // the page — the shape of every "the harness bypassed it" bug in this repo.
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
  return document.getElementById("pane") as HTMLElement;
}

function pane(): HTMLElement {
  return document.getElementById("pane") as HTMLElement;
}
/**
 * Where THIS deck's crash record lives.
 *
 * One key per deck: a single key made every write a choice between destroying
 * another deck's record of slides still sitting in it and denying this deck a
 * record at all. The literal is repeated rather than shared with the mock
 * above, because `vi.mock` is hoisted above every `const` in this file.
 */
const CRUMB_KEY = "ssf-merge.run.v1:https://example-my.sharepoint.com/personal/x/Documents/deck.pptx";

function primary(): HTMLButtonElement {
  return pane().querySelector("button.primary") as HTMLButtonElement;
}
/**
 * The card's plain "Remove the preview", which takes the row out and STAYS.
 *
 * The primary beside it removes and carries on to the merge, so the two are
 * different journeys and every test below has to name which one it means.
 */
function removePreview(): HTMLButtonElement {
  return pane().querySelector('[data-action="end-preview"]') as HTMLButtonElement;
}
function field(name: string): HTMLInputElement | HTMLTextAreaElement {
  return pane().querySelector(`[data-field="${name}"]`) as HTMLInputElement;
}
function type(name: string, value: string): void {
  const node = field(name);
  node.value = value;
  node.dispatchEvent(new Event("input", { bubbles: true }));
}
/** Every sentence the pane is showing the user right now. */
function said(): string[] {
  return Array.from(pane().querySelectorAll(".blocked, .notice")).map((n) => n.textContent ?? "");
}
/** Let queued microtasks and any awaited promise chain settle. */
async function settle(): Promise<void> {
  for (let i = 0; i < 8; i++) await Promise.resolve();
}

/**
 * Walk from a fresh pane to the merge step with a block and data in hand.
 *
 * Five steps: template, data, fields, preview, merge. TWO template reads, not
 * one — the fields step re-reads the slides because the user has just been
 * putting `{{Column}}` onto them and nothing tells the pane that happened.
 */
async function reachMerge(): Promise<void> {
  await openPane();
  await settle();
  type("from", "4");
  type("to", "6");
  office.inspectBlock.mockResolvedValueOnce(REPORT);
  primary().click(); // template -> data
  await settle();
  type("paste", "First\tLast\nAda\tLovelace\nGrace\tHopper");
  primary().click(); // data -> fields
  office.inspectBlock.mockResolvedValueOnce(REPORT);
  primary().click(); // fields, re-reading the slides -> preview
  await settle();
  // The preview step's primary SHOWS a row rather than advancing, so the way
  // forward is the link beside it.
  (pane().querySelector("[data-forward]") as HTMLElement).click();
}

beforeEach(() => {
  // The crash crumb lives in `localStorage`, which `vi.resetModules` does not
  // touch — so a merge in one test was read back as a recovered run by the
  // next, which now draws its offer on whatever step the pane is on. Cleared
  // here rather than in the one describe that knew about it: every test in
  // this file opens a pane, and a pane reads the crumb.
  localStorage.clear();
  office.slideCount.mockReset().mockResolvedValue(12);
  office.inspectBlock.mockReset();
  office.runMerge.mockReset();
  office.undoMerge.mockReset();
  office.selectedBlock.mockReset();
  office.canReadSelection.mockReset().mockReturnValue(true);
  office.insertTextAtCursor.mockReset().mockResolvedValue({ ok: true });
  office.ready.mockReturnValue({ ok: true, detail: "fine" });
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("the boot crumb read", () => {
  it("does not speak about, or delete, a run that is happening right now", async () => {
    /**
     * The deck count this handler waits on is issued at boot and can answer
     * minutes later — long enough for the user to walk the wizard and press
     * Merge. The crumb it then found was the pending marker THAT run had just
     * written, so the pane said the merge "did not finish" while it was visibly
     * running, and deleted the record in the one window it exists for.
     */
    let count: (n: number) => void = () => undefined;
    office.slideCount.mockReset().mockReturnValueOnce(new Promise((res) => (count = res)));
    await openPane();
    await settle();

    // The boot count is still out. The user walks the wizard and merges.
    office.slideCount.mockResolvedValue(12);
    type("from", "4");
    type("to", "6");
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click();
    await settle();
    type("paste", "First\tLast\nAda\tLovelace");
    primary().click();
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click();
    await settle();
    (pane().querySelector("[data-forward]") as HTMLElement).click();
    await settle();

    let landed: (o: unknown) => void = () => undefined;
    office.runMerge.mockReturnValueOnce(new Promise((res) => (landed = res)));
    primary().click();
    await settle();
    const during = localStorage.getItem(CRUMB_KEY);
    expect(during, "the run wrote its pending marker").not.toBeNull();

    // Now the boot count finally answers, mid-merge.
    count(12);
    await settle();

    expect(document.body.textContent, "about a merge that is running").not.toContain("did not finish");
    expect(localStorage.getItem(CRUMB_KEY), "the record the run may need").not.toBeNull();

    landed({ ...OUTCOME, deckAtStart: 12 });
    await settle();
  });
});

it("still reports a genuine crash record while an unrelated read is running", async () => {
  /**
   * The first guard was `!state.running`, which is too wide: this branch is
   * never retried, so ANY run in flight when the boot count lands — the
   * "Reading the slides…" a user is most likely to start first — swallowed a
   * real crash record for the whole session.
   *
   * The marker's id is unique per run now, so the pane recognises its OWN and
   * says nothing only about that.
   */
  localStorage.setItem(
    CRUMB_KEY,
    JSON.stringify({
      kind: "ssf-merge-run",
      deckAtStart: 12,
      added: 0,
      runId: "died-in-another-session",
      startedAt: "2026-08-27T10:00:00.000Z",
      doc: "https://example-my.sharepoint.com/personal/x/Documents/deck.pptx",
    }),
  );
  let count: (n: number) => void = () => undefined;
  office.slideCount.mockReset().mockReturnValueOnce(new Promise((res) => (count = res)));
  await openPane();
  await settle();

  // An ordinary template read is in flight when the boot count answers.
  office.slideCount.mockResolvedValue(12);
  type("from", "4");
  type("to", "6");
  let read: (r: unknown) => void = () => undefined;
  office.inspectBlock.mockReturnValueOnce(new Promise((res) => (read = res)));
  primary().click();
  await settle();
  count(18);
  await settle();

  expect(document.body.textContent, "a crash record this pane did not write").toContain("did not finish");
  read(REPORT);
  await settle();
});

describe("a read that answers after the user has moved", () => {
  it("leaves the user on the step they walked to", async () => {
    /**
     * The template read takes seconds and the Back link stays live throughout —
     * deliberately, because a user who realises they typed the wrong slides
     * should not be held. The advance at the end of the read was unconditional,
     * so walking back to Data while it was running jumped the pane forward to
     * Preview when the answer arrived, over the work the user had gone back to
     * change.
     *
     * The block's own staleness was already re-checked. Where the USER is was
     * not.
     */
    await openPane();
    await settle();
    type("from", "4");
    type("to", "6");
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // template -> data
    await settle();
    type("paste", "First\tLast\nAda\tLovelace");
    primary().click(); // data -> fields
    await settle();
    expect(pane().querySelector(".step-of")?.textContent).toBe("Step 3 of 5 · Fields");

    // The fields step re-reads the slides, because the user has just been
    // typing `{{Column}}` into PowerPoint and nothing tells the pane so.
    let answer: (r: unknown) => void = () => undefined;
    office.inspectBlock.mockReturnValueOnce(new Promise((res) => (answer = res)));
    primary().click(); // fields -> reading
    await settle();

    // The user changes their mind about the data and walks back while it reads.
    (pane().querySelector("[data-back]") as HTMLElement).click();
    await settle();
    const walkedTo = pane().querySelector(".step-of")?.textContent;
    expect(walkedTo, "the user moved").toBe("Step 2 of 5 · Data");

    answer(REPORT);
    await settle();
    expect(pane().querySelector(".step-of")?.textContent, "the read moved the user").toBe(walkedTo);
  });
});

describe("a merge that raises", () => {
  it("says so, and gives the button back", async () => {
    // The committed version awaited runMerge with no catch, so a raise left the
    // hand-written `disabled = true` / "Merging…" on a button no later draw
    // ever replaced, showed the user nothing at all, and lost the numbers an
    // undo is clamped against. `readTemplate` throws its refusal by design, so
    // this is the ordinary path, not an exotic one.
    await reachMerge();
    expect(primary().textContent).toBe("Add 6 slides");
    office.runMerge.mockRejectedValueOnce(new Error(readTemplateRefusal));
    primary().click();
    await settle();

    expect(said().join(" ")).toContain(readTemplateRefusal);
    expect(primary().disabled, "the button comes back").toBe(false);
    expect(primary().textContent).not.toBe("Merging…");
  });

  it("counts the deck again, because a raise does not mean nothing happened", async () => {
    // This host takes calls it does not perform AND performs calls it then
    // raises on: insertDeck's own confirming re-count sits outside the try it
    // uses to read the delta, so a timeout there rejects with the slides
    // already in the deck. The DELTA is the evidence, never the absence of an
    // error.
    await reachMerge();
    // Two counts now, and the FIRST is the floor: taken just before the insert
    // rather than read off the pane's cache. See `deckBefore` in `main.ts`.
    office.slideCount.mockResolvedValueOnce(12).mockResolvedValueOnce(18);
    office.runMerge.mockRejectedValueOnce(new Error("gave up waiting for: inserting the merged deck"));
    primary().click();
    await settle();

    const told = said().join(" ");
    expect(told).toContain("6 slide");
    expect(told).toContain("landed anyway");
  });
});

describe("what a screen reader hears while the pane works", () => {
  /**
   * The pure decision is in `pane-steps.test.ts`. This is the wiring: that a
   * live region exists at all, that it is made ONCE rather than rebuilt — a
   * region created with its content already in it does not announce — and that
   * the real `draw()` writes into it.
   *
   * The pane's markup here is the same two elements `taskpane.html` has, and
   * deliberately does NOT include the region: it is created by `main.ts`, so a
   * page that never got one still announces.
   */
  const announcer = () => document.getElementById("announcer");

  it("makes one live region, correctly marked, and keeps it", async () => {
    await reachMerge();
    const first = announcer();
    expect(first, "no live region was made").not.toBeNull();
    expect(first?.getAttribute("aria-live")).toBe("polite");
    expect(first?.getAttribute("role")).toBe("status");
    // Outside the pane, because `render` empties the pane on every draw.
    expect(pane().contains(first)).toBe(false);

    office.runMerge.mockResolvedValueOnce(OUTCOME);
    primary().click();
    await settle();
    // The SAME node across many draws. A fresh one per draw is a region that
    // announces nothing.
    expect(announcer()).toBe(first);
    expect(document.querySelectorAll("#announcer")).toHaveLength(1);
  });

  it("says what happened when the merge is over", async () => {
    await reachMerge();
    expect(announcer()?.textContent).toBe("");
    office.runMerge.mockResolvedValueOnce(OUTCOME);
    primary().click();
    await settle();
    expect(announcer()?.textContent).toContain("6 slides added after slide 12.");
  });

  it("says what is out while a call is out, then what it did", async () => {
    // The window this whole thing exists for: a merge is legitimately silent
    // for minutes, and the difference between slow and stuck is the only
    // question anybody has.
    //
    // The mock stands in for the engine, so it has to emit the host trace the
    // real `runMerge` emits — that trace is what sets `inFlight`, through the
    // subscription `merge()` opens. Without it this would assert against a
    // field nothing in the test ever sets.
    await reachMerge();
    // AFTER the pane is open. `openPane` calls `vi.resetModules()` and then
    // imports `main.ts`, so a `trace` imported before that is a different
    // module instance from the one the pane subscribed to — the subscription
    // is module state, and this asserted against an empty region for exactly
    // that reason before the import moved.
    const { trace } = await import("../src/core/trace.js");
    const held = deferred<unknown>();
    office.runMerge.mockImplementationOnce(() => {
      trace("host", "issued", { call: "inserting the merged deck" });
      return held.promise;
    });
    primary().click();
    await settle();
    expect(announcer()?.textContent).toBe("Waiting on PowerPoint: inserting the merged deck");

    held.resolve(OUTCOME);
    await settle();
    expect(announcer()?.textContent).toContain("6 slides added after slide 12.");
  });

  it("goes quiet when the sentence it was reading is gone", async () => {
    // Walking back clears the notice, so there is nothing to say any more. A
    // live region left holding the last run's outcome is a pane telling a
    // screen reader user about a screen they are no longer on.
    await reachMerge();
    office.runMerge.mockResolvedValueOnce(OUTCOME);
    primary().click();
    await settle();
    expect(announcer()?.textContent).not.toBe("");
    (pane().querySelector("[data-back]") as HTMLElement).click();
    await settle();
    expect(announcer()?.textContent).toBe("");
  });

  it("does not re-announce the same sentence on a redraw", async () => {
    // The pane redraws on every keystroke. Writing the same string back into a
    // live region makes some screen readers say it again, so the write is gated
    // on a change — watched here on the node itself.
    //
    // A DISCLOSURE toggle is the redraw to use: it changes the screen and
    // deliberately leaves `notice` alone. Walking back was the first attempt
    // and is the wrong control — it clears the notice, so the region correctly
    // changes, which the test above now asserts instead.
    await reachMerge();
    office.runMerge.mockResolvedValueOnce(OUTCOME);
    primary().click();
    await settle();
    const node = announcer() as HTMLElement;
    const said = node.textContent;
    let writes = 0;
    const observer = new MutationObserver(() => writes++);
    observer.observe(node, { childList: true, characterData: true, subtree: true });
    (pane().querySelector('[data-action="rows"]') as HTMLElement).click();
    (pane().querySelector('[data-action="rows"]') as HTMLElement).click();
    await settle();
    observer.disconnect();
    expect(writes, "the live region was written again with the same text").toBe(0);
    expect(node.textContent, "and it still holds what it said").toBe(said);
  });
});

describe("a merge that is still running", () => {
  it("cannot be started twice by going back and forward", async () => {
    // The only thing stopping a re-press was `button.disabled = true` on a DOM
    // node, and every later draw replaced it with one `primary()` had
    // re-enabled. Back then Continue during a two-minute merge gave the user a
    // live "Add 6 slides" over a run already in flight, and their deck got both
    // inserts with only one set of undo clamps surviving.
    await reachMerge();
    const held = deferred<unknown>();
    office.runMerge.mockReturnValueOnce(held.promise);
    primary().click();
    await settle();
    expect(primary().disabled).toBe(true);

    (pane().querySelector("[data-back]") as HTMLElement).click(); // to preview
    primary().click(); // forward to merge again
    await settle();
    expect(primary().disabled, "still frozen after a round trip").toBe(true);
    primary().click();
    await settle();
    expect(office.runMerge).toHaveBeenCalledTimes(1);

    held.resolve(OUTCOME);
    await settle();
  });

  it("does not offer to add the same slides again once they are added", async () => {
    // The merge screen redraws after a successful run, and a live "Add 6
    // slides" beside a notice saying six slides were added is how a deck gets
    // twelve.
    await reachMerge();
    office.runMerge.mockResolvedValueOnce(OUTCOME);
    primary().click();
    await settle();
    expect(primary().textContent).toBe("Added 6 slides");
    expect(primary().disabled).toBe(true);

    // Pressing it again adds nothing.
    primary().click();
    await settle();
    expect(office.runMerge).toHaveBeenCalledTimes(1);

    // An edit is what re-arms it, because an edit is a different merge. Back
    // to the paste box, change the data, and the button is live again.
    (pane().querySelector("[data-back]") as HTMLElement).click(); // preview
    (pane().querySelector("[data-back]") as HTMLElement).click(); // fields
    (pane().querySelector("[data-back]") as HTMLElement).click(); // data
    type("paste", "First\tLast\nAda\tLovelace");
    primary().click(); // data -> fields
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // fields, re-reading the slides -> preview
    await settle();
    (pane().querySelector("[data-forward]") as HTMLElement).click(); // -> merge
    expect(primary().textContent).toBe("Add 3 slides");
    expect(primary().disabled).toBe(false);
  });
});

describe("a template read that is still running", () => {
  it("freezes its own button rather than firing a second export", async () => {
    await openPane();
    await settle();
    type("from", "4");
    type("to", "6");
    const held = deferred<unknown>();
    office.inspectBlock.mockReturnValueOnce(held.promise);
    primary().click();
    await settle();
    expect(primary().textContent).toBe("Reading the slides…");
    expect(primary().disabled).toBe(true);
    primary().click();
    await settle();
    expect(office.inspectBlock).toHaveBeenCalledTimes(1);
    held.resolve(REPORT);
    await settle();
  });

  it("refuses to commit a block the boxes no longer name", async () => {
    // The read captures the block before the await and wrote it back after.
    // Retyping mid-read left state.fields describing slides 4 to 6 while
    // chosenBlock answered 4 to 9 — and the merge runs on chosenBlock, so it
    // would have cloned slides nobody read and bound one block's placeholders
    // to another block's slides.
    await openPane();
    await settle();
    type("from", "4");
    type("to", "6");
    const held = deferred<unknown>();
    office.inspectBlock.mockReturnValueOnce(held.promise);
    primary().click();
    await settle();

    // The keystroke is refused while the read is out, so the boxes cannot
    // drift out from under the answer that is on its way back.
    type("to", "9");
    await settle();
    held.resolve(REPORT);
    await settle();

    // The read committed slides 4 to 6 and moved on, and the boxes still say
    // 4 and 6 — not 4 and 9 with 4-to-6's placeholders behind them.
    expect(office.inspectBlock).toHaveBeenCalledWith({ from: 4, to: 6 });
    (pane().querySelector("[data-back]") as HTMLElement).click();
    expect((field("from") as HTMLInputElement).value).toBe("4");
    expect((field("to") as HTMLInputElement).value).toBe("6");
    expect(primary().textContent).toBe("Use slides 4 to 6");
  });

  it("says so when the read raises, instead of reading forever", async () => {
    await openPane();
    await settle();
    type("from", "4");
    type("to", "6");
    office.inspectBlock.mockRejectedValueOnce(new Error("the package has no part"));
    primary().click();
    await settle();
    expect(said().join(" ")).toContain("the package has no part");
    expect(said().join(" ")).not.toContain("Reading the slides…");
    expect(primary().disabled).toBe(false);
  });
});

describe("the caret", () => {
  it("stays where the user put it, across the redraw every keystroke causes", async () => {
    // render() empties the root and builds fresh elements, so every draw
    // destroys the focused box. The first version focused the replacement and
    // sent the caret to the END — so typing 5 into "4|6" gave "456" and the
    // next digit landed after the 6. In the paste box it scattered the rest of
    // a line to the end and readPastedTable merged the corrupted text.
    await openPane();
    await settle();
    const box = field("from") as HTMLInputElement;
    box.value = "46";
    box.focus();
    box.setSelectionRange(1, 1);
    box.value = "456";
    box.setSelectionRange(2, 2);
    box.dispatchEvent(new Event("input", { bubbles: true }));

    const after = field("from") as HTMLInputElement;
    expect(document.activeElement).toBe(after);
    expect(after.selectionStart).toBe(2);
    expect(after.selectionEnd).toBe(2);
  });

  it("survives a redraw the user did not cause", async () => {
    // The deck count resolves a second or two after the pane opens and
    // redraws. That redraw blanked the focus, so the next digit typed went to
    // <body> and vanished — with the box still holding what looked like a
    // dropped keystroke.
    const count = deferred<number>();
    office.slideCount.mockReturnValueOnce(count.promise);
    await openPane();
    const box = field("from") as HTMLInputElement;
    box.focus();
    box.value = "4";
    box.dispatchEvent(new Event("input", { bubbles: true }));
    expect(document.activeElement).toBe(field("from"));

    count.resolve(12);
    await settle();
    expect(document.activeElement, "still in the box the user is typing in").toBe(field("from"));
  });

  it("stays on the control that was pressed, when that control is still there", async () => {
    // The caret rule was written for the two text boxes and nothing else, so
    // every OTHER control lost focus on its own press: the redraw rebuilds the
    // element and `draw` only looked for `data-field`. A keyboard user
    // unticking rows in a 200-row list was thrown to the top of the pane after
    // each one, and had to tab back down to reach the next.
    //
    // Four controls, each pressed and each still on the screen afterwards.
    await reachMerge();

    const rows = document.querySelector<HTMLButtonElement>('[data-action="rows"]');
    rows?.focus();
    rows?.click();
    await settle();
    expect(document.activeElement, "the disclosure that just opened the list").toBe(
      document.querySelector('[data-action="rows"]'),
    );

    const box = document.querySelector<HTMLInputElement>('input[data-row="0"]');
    box?.focus();
    box?.click();
    await settle();
    expect(document.activeElement, "the row that was just unticked").toBe(
      document.querySelector('input[data-row="0"]'),
    );

    const conditions = document.querySelector<HTMLButtonElement>('[data-action="conditions"]');
    conditions?.focus();
    conditions?.click();
    await settle();
    expect(document.activeElement, "the disclosure that just opened the conditions").toBe(
      document.querySelector('[data-action="conditions"]'),
    );

    // A condition select had its own branch before this and no test. It is one
    // of the same set now, and the reason it was singled out holds: this is a
    // control whose whole use is setting several in a row.
    const select = document.querySelector<HTMLSelectElement>('[data-condition="5"]');
    if (select) {
      select.focus();
      select.value = "First";
      select.dispatchEvent(new Event("input", { bubbles: true }));
      await settle();
      expect(document.activeElement, "the condition just chosen").toBe(document.querySelector('[data-condition="5"]'));
    }
    expect(select, "the conditions were open at all").not.toBeNull();

    // And a chip, which is on another step. `insertTextAtCursor` resolves, the
    // chips are rebuilt with a new `data-placed`, and `data-insert` is what
    // says it is the same chip.
    (pane().querySelector("[data-back]") as HTMLElement).click(); // preview
    (pane().querySelector("[data-back]") as HTMLElement).click(); // fields
    const chip = document.querySelector<HTMLButtonElement>('[data-insert="First"]');
    chip?.focus();
    chip?.click();
    await settle();
    expect(document.activeElement, "the chip that was just pressed").toBe(
      document.querySelector('[data-insert="First"]'),
    );
  });

  it("uses text boxes, because a number input will not say where the caret is", async () => {
    // type="number" answers null for selectionStart and throws on
    // setSelectionRange, so the caret cannot be restored across a redraw at
    // all. inputmode keeps the digit keyboard; readBlockDraft was always the
    // authority on what the box holds.
    await openPane();
    await settle();
    for (const name of ["from", "to"]) {
      expect(field(name).getAttribute("type"), name).toBe("text");
      expect(field(name).getAttribute("inputmode"), name).toBe("numeric");
    }
  });
});

describe("the preview", () => {
  /** Walk to the preview step with a block and two rows in hand. */
  async function reachPreview(): Promise<void> {
    await openPane();
    await settle();
    type("from", "4");
    type("to", "6");
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // template -> data
    await settle();
    type("paste", "First\tLast\nAda\tLovelace\nGrace\tHopper");
    primary().click(); // data -> fields
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // fields, re-reading the slides -> preview
    await settle();
  }

  const PREVIEW = { ...OUTCOME, added: 3, detail: "3 slides added after slide 12." };

  it("names the slides the preview actually landed on", async () => {
    /**
     * The card exists so a user who closes the pane can find the preview slides
     * and delete them by hand, so the numbers on it have to be the numbers on
     * the rail. It took them from `deckAtStart` — the deck's size when the run
     * was PLANNED — and a slide arriving between then and the insert moves
     * every one of them by one: the card named the co-author's slide and left
     * the last of the preview's own out.
     *
     * The merge summary was given this anchor in the same round; its sibling
     * one function up was not.
     */
    await reachPreview();
    office.runMerge.mockResolvedValueOnce({ ...PREVIEW, landedAfter: 13 });
    primary().click();
    await settle();

    expect(pane().textContent).toContain("Slides 14 to 16 are a preview");
    expect(pane().textContent, "the slide a co-author added is not the preview's").not.toContain("Slides 13 to 15");
  });

  it("runs the ORDINARY merge over one row", async () => {
    // The whole value of the step. A preview rendered by some other route is a
    // preview of something nobody is going to get — and writing the row onto
    // the template through Office.js, which is what the backlog specified,
    // re-authors the text (office-js#5858) on the one slide the product exists
    // to preserve.
    await reachPreview();
    office.runMerge.mockResolvedValueOnce(PREVIEW);
    primary().click();
    await settle();

    expect(office.runMerge).toHaveBeenCalledTimes(1);
    const req = office.runMerge.mock.calls[0]?.[0] as { records: { rows: unknown[] }; from: number; to: number };
    expect(req.records.rows, "one row, not all of them").toHaveLength(1);
    expect(req.from).toBe(4);
    expect(req.to).toBe(6);
  });

  it("names the slides it landed on", async () => {
    await reachPreview();
    office.runMerge.mockResolvedValueOnce(PREVIEW);
    primary().click();
    await settle();
    // deckAtStart 12, added 3 — so slides 13 to 15.
    expect(pane().querySelector(".card.undo")?.textContent).toContain("Slides 13 to 15");
    expect(primary().textContent).toBe("On to the merge");
    expect(removePreview().textContent, "the plain way back is on the card").toBe("Remove the preview");
  });

  it("takes it back with the same clamped sweep an undo uses", async () => {
    await reachPreview();
    office.runMerge.mockResolvedValueOnce(PREVIEW);
    primary().click();
    await settle();

    office.undoMerge.mockResolvedValueOnce({ removed: 3, detail: "removed 3 slide(s) from index 12" });
    removePreview().click();
    await settle();

    expect(office.undoMerge).toHaveBeenCalledWith(expect.objectContaining({ deckAtStart: 12, added: 3 }));
    expect(pane().querySelectorAll(".card.undo")).toHaveLength(0);
    expect(primary().textContent).toBe("Preview the first row");
  });

  it("carries ON to the merge in ONE press, taking the preview out on the way", async () => {
    // The journey the step exists for, and the one it used to make people
    // guess at. While a preview was up the pane offered "Back to fields" and
    // "Remove the preview" and nothing else — the word "merge" was nowhere on
    // the screen — so the way on was to work out that clearing up was the way
    // on, and it took four presses.
    await reachPreview();
    office.runMerge.mockResolvedValueOnce(PREVIEW);
    primary().click();
    await settle();
    expect(pane().textContent, "the merge has to be nameable from here").toContain("merge");

    office.undoMerge.mockResolvedValueOnce({ removed: 3, detail: "removed 3 slide(s) from index 12" });
    primary().click();
    await settle();

    // The preview is gone AND the wizard has moved on, off the one press.
    expect(office.undoMerge).toHaveBeenCalledWith(expect.objectContaining({ deckAtStart: 12, added: 3 }));
    expect(pane().querySelector(".step-of")?.textContent).toBe("Step 5 of 5 · Merge");
    expect(pane().querySelectorAll(".card.undo")).toHaveLength(0);
  });

  it("stays put when the sweep could not take the whole preview back", async () => {
    // Advancing on a partial removal would land the user on a merge step
    // refusing with "End the preview before merging." while the sentence
    // explaining what actually happened sat one screen behind them.
    await reachPreview();
    office.runMerge.mockResolvedValueOnce(PREVIEW);
    primary().click();
    await settle();

    office.undoMerge.mockResolvedValueOnce({ removed: 1, detail: "removed 1 slide(s) from slides 13 to 13" });
    // The DECK says so too: three went in, one came out, so two are still
    // there. The pane asks before it claims, because the commonest way to
    // reach a partial removal is that the user deleted the slides themselves —
    // and then nothing is left behind and the deck is back where it started.
    office.slideCount.mockResolvedValueOnce(14);
    primary().click();
    await settle();

    expect(pane().querySelector(".step-of")?.textContent).toBe("Step 4 of 5 · Preview");
    expect(pane().textContent).toContain("Some of the preview is still there");
    // And it stops NAMING the slides. The ones that went took the numbering of
    // the ones that stayed with them, so "Slides 12 to 14 are a preview of the
    // first row" was on screen over a deck where 12 is the user's own slide
    // again, beside a button offering to delete them.
    expect(pane().textContent, "a range the deck no longer answers to").not.toMatch(/Slides? \d+.*are a preview/);
    expect(pane().textContent).toContain("A preview is in your deck.");
  });

  it("does not hold the user on the preview over a slide the sweep DISOWNED", async () => {
    /**
     * The merge undo subtracts `disowned` — a slide the sweep declined because
     * it carries no mark of this run is not one the pane is waiting to take
     * back — and this screen did not. Preview three slides, delete one of them
     * by hand and append one of your own: two come back, one is disowned, and
     * the pane said "Some of the preview is still there" about the slide it had
     * just called not the preview's, with `previewing` left set.
     *
     * That is terminal. While a preview is up the forward link is withheld and
     * the merge step refuses, so the user could not reach the merge again for
     * the rest of the session, over a slide nothing was ever going to remove.
     */
    await reachPreview();
    office.runMerge.mockResolvedValueOnce(PREVIEW);
    primary().click();
    await settle();

    office.undoMerge.mockResolvedValueOnce({
      removed: 2,
      disowned: 1,
      detail: "removed 2 slide(s); 1 was not this run's",
    });
    // The deck says the disowned slide is still there, so the "already gone"
    // branch is not what carries this — without the `disowned` subtraction the
    // pane takes the partial branch and stays put.
    office.slideCount.mockResolvedValueOnce(13);
    primary().click();
    await settle();

    expect(pane().textContent, "about a slide the sweep called not the preview's").not.toContain("still there");
    expect(pane().querySelector(".step-of")?.textContent, "the way on").toBe("Step 5 of 5 · Merge");
  });

  it("never holds the wizard on the preview after a press that changed nothing", async () => {
    /**
     * The INVARIANT, over the whole matrix rather than the three cases that
     * were reported one at a time. This screen has produced the same terminal
     * state three times by three routes — a slide the sweep disowned, a count
     * carried forward that `sweepPlan` then refused, and a deck a co-author
     * grew — and each was found only after somebody hit it.
     *
     * The rule is not "always leave the step". Where a press moved something
     * the next press has less to do and can finish, and taking the way back
     * away there would strand a preview that was halfway out. The rule is that
     * `previewing` may only stay set when the press CHANGED something, because
     * that is the only case in which pressing again is worth offering.
     */
    for (const removed of [0, 1, 2, 3]) {
      for (const disowned of [0, 1]) {
        if (removed + disowned > 3) continue;
        for (const deckNow of [9, 12, 13, 15, 17]) {
          await reachPreview();
          office.runMerge.mockResolvedValueOnce(PREVIEW);
          primary().click();
          await settle();

          office.undoMerge.mockResolvedValueOnce({ removed, disowned, detail: `removed ${removed}` });
          office.slideCount.mockResolvedValueOnce(deckNow);
          primary().click();
          await settle();

          const where = `removed=${removed} disowned=${disowned} deckNow=${deckNow}`;
          const stillOnPreview = pane().querySelector(".step-of")?.textContent === "Step 4 of 5 · Preview";
          if (stillOnPreview) {
            expect(removed + disowned, `${where}: held on the preview by a press that did nothing`).toBeGreaterThan(0);
          }
        }
      }
    }
  });

  it("asks the sweep for PROOF on a second press", async () => {
    // The pane's half of the rule: a press that leaves slides owed carries the
    // outcome forward marked `pressed`, and `undoMerge` turns that into
    // `requireProof` — so the sweep may no longer fall back to position on a
    // deck that has provably changed shape since the run. Without the mark, a
    // host that stops answering tags between presses takes the whole window,
    // and a slide the user made in the meantime goes with it.
    await reachPreview();
    office.runMerge.mockResolvedValueOnce(PREVIEW);
    primary().click();
    await settle();

    office.undoMerge.mockResolvedValueOnce({ removed: 1, disowned: 0, detail: "removed 1 slide(s)" });
    office.slideCount.mockResolvedValueOnce(14);
    primary().click();
    await settle();

    office.undoMerge.mockResolvedValueOnce({ removed: 2, disowned: 0, detail: "removed 2 slide(s)" });
    primary().click();
    await settle();
    expect(office.undoMerge.mock.calls[1]?.[0], "the second press says it is one").toMatchObject({ pressed: true });
    expect(office.undoMerge.mock.calls[0]?.[0], "and the first says it is not").not.toMatchObject({ pressed: true });
  });

  it("stops offering the preview back once a press has DECLINED a slide", async () => {
    /**
     * The pane may not press again after the sweep has met a slide it will not
     * claim. Carrying a count forward re-submits that slide to a window it is
     * now alone in, where an all-untagged plan is taken whole — so the second
     * press deletes the user's own slides, under a notice saying there was
     * nothing to take back. `test/undo.test.ts` walks that mechanism in the
     * pure code; this holds the pane to the decision.
     */
    await reachPreview();
    office.runMerge.mockResolvedValueOnce(PREVIEW);
    primary().click();
    await settle();

    // One out, one declined, of three: something of this run's is still owed,
    // so this is not the "nothing left" path — it is the one where a second
    // press would reach across the declined slide.
    office.undoMerge.mockResolvedValueOnce({ removed: 1, disowned: 1, detail: "removed 1 slide(s)" });
    office.slideCount.mockResolvedValueOnce(14);
    const before = office.undoMerge.mock.calls.length;
    primary().click();
    await settle();

    expect(pane().querySelector(".step-of")?.textContent, "a pane must always leave a way on").toBe(
      "Step 5 of 5 · Merge",
    );
    // COULD NOT BE SHOWN, not "are not this run's": a slide with no tag and a
    // slide the host would not answer for arrive here as the same `undefined`,
    // and on a host below PowerPointApi 1.3 the stronger sentence is false for
    // every user over slides that ARE the run's.
    expect(pane().textContent).toContain("could not be shown to be this run's");
    expect(pane().textContent, "a claim the data cannot support").not.toContain("are not this run's");
    // COUNTED, not "the rest". One slide was declined and one is still owed,
    // and the sentence spoke for both — the same defect the merge undo's
    // notice carried, in the sibling screen that was not changed with it.
    expect(pane().textContent).toContain("1 slide could not be shown to be this run's");
    expect(pane().textContent, "a claim about slides the sweep never doubted").not.toContain("The rest");
    // And there is no second press to make: the preview is over.
    expect(office.undoMerge.mock.calls.length, "no further sweep is offered").toBe(before + 1);
  });

  it("lets the user out when a press takes nothing back at all", async () => {
    /**
     * A co-author adds a slide during the preview: the deck has grown by more
     * than the run added, and `sweepPlan` refuses the shape because it can no
     * longer prove which slides are the preview's. The press removes nothing
     * and disowns nothing, and the next press will answer the same way — so
     * holding the user on the preview step holds them there for the session,
     * with the forward link withheld and the merge step refusing.
     *
     * That is the terminal state this screen has now produced three times by
     * three routes. A press that MOVED something still keeps the way back, and
     * the difference between the two is the whole of the fix: "press again" is
     * worth offering only when pressing again can do something.
     */
    await reachPreview();
    office.runMerge.mockResolvedValueOnce(PREVIEW);
    primary().click();
    await settle();

    office.undoMerge.mockResolvedValueOnce({
      removed: 0,
      disowned: 0,
      detail: "nothing to take back (the deck grew by 5 while this run added 3)",
    });
    // The deck is BIGGER than it started, so "already gone" does not carry it.
    office.slideCount.mockResolvedValueOnce(17);
    primary().click();
    await settle();

    expect(pane().querySelector(".step-of")?.textContent, "a pane must always leave a way on").toBe(
      "Step 5 of 5 · Merge",
    );
    expect(pane().textContent).toContain("could not be taken back");
    expect(pane().textContent, "and what to do about the slides").toContain("can be deleted from the thumbnail rail");
  });

  it("can still end a preview after a press that left some of it behind", async () => {
    /**
     * `sweepPlan` clamps `added` against the deck's GROWTH, so the count
     * carried forward after a partial press has to be a deck size. Subtracting
     * the DISOWNED slide as well — which is right for deciding what to say —
     * made `grew` exceed `added` on the next press, so the plan came back null
     * forever, on the one branch that does not clear `previewing`.
     *
     * That is terminal: the forward link is withheld, the rail is pinned and
     * the merge step refuses while a preview is up. The fix for a deadlock
     * introduced a worse one, and the failing shape is `removed + disowned <
     * added` with `disowned > 0`, which the test written beside it did not have.
     */
    await reachPreview();
    office.runMerge.mockResolvedValueOnce(PREVIEW);
    primary().click();
    await settle();

    // One removed of three, and NONE declined: the window still holds only
    // this run's slides, so a second press is worth offering. A press that
    // declined one would end the offer instead — see the test below.
    office.undoMerge.mockResolvedValueOnce({ removed: 1, disowned: 0, detail: "removed 1 slide(s)" });
    office.slideCount.mockResolvedValueOnce(14);
    primary().click();
    await settle();
    expect(pane().textContent, "one is still outstanding").toContain("Some of the preview is still there");

    // The second press must ask for what the DECK is holding — three added
    // less the one that went — because `sweepPlan` clamps that number against
    // the deck's growth and the disowned slide is still in the deck. Asked for
    // two, the plan is null and the press does nothing, forever.
    office.undoMerge.mockResolvedValueOnce({ removed: 2, disowned: 0, detail: "removed 2 slide(s)" });
    primary().click();
    await settle();
    const second = office.undoMerge.mock.calls[1]?.[0] as { added: number; deckAtStart: number };
    expect(second.added, "a count the sweep is clamped against is a deck size").toBe(2);
    expect(second.deckAtStart, "and the floor it is measured from does not move").toBe(PREVIEW.deckAtStart);
    expect(pane().querySelector(".step-of")?.textContent, "the way on").toBe("Step 5 of 5 · Merge");
  });

  it("carries on when the slides the sweep could not take are already gone", async () => {
    /**
     * The commonest way to reach a partial removal: the user did what the card
     * says the button does and deleted the preview slides themselves. The deck
     * never grew, the sweep refused, and this said "Some of the preview is
     * still there — nothing to take back (deck was 12, is 12)" — one clause
     * contradicting the next.
     *
     * And it returned with `previewing` still set, which was TERMINAL. While a
     * preview is up the forward link is withheld, the rail is not clickable and
     * the merge step refuses, so the user could not reach the merge again for
     * the rest of the session, with nothing on screen saying why.
     */
    await reachPreview();
    office.runMerge.mockResolvedValueOnce(PREVIEW);
    primary().click();
    await settle();

    office.undoMerge.mockResolvedValueOnce({ removed: 0, detail: "nothing to take back (deck was 12, is 12)" });
    office.slideCount.mockResolvedValueOnce(12);
    primary().click();
    await settle();

    // SIZE, never identity. "Those slides are already gone" was a claim about
    // WHICH slides made from a count of them, and a deck back to the size it
    // started at is equally the user having deleted their OWN slides with
    // preview slides still in it. The sentence says what can be done instead of
    // what is there.
    expect(pane().textContent).toContain("could not be taken back");
    expect(pane().textContent, "a claim of identity taken from a size").not.toContain("already gone");
    expect(pane().textContent, "nor the same claim by another spelling").not.toContain("nothing here to take back");
    // "Your deck holds 12 slides, no more than before the preview" was here for
    // one commit and is the same claim in the same place: these exact inputs
    // are BOTH worlds, and it suppressed the advice in the one where the
    // preview slides are still in the deck. The count is reported; the advice
    // is unconditional.
    expect(pane().textContent, "a size read as an identity").not.toContain("no more than before the preview");
    expect(pane().textContent).toContain("Your deck holds 12 slides");
    expect(pane().textContent).toContain("can be deleted from the thumbnail rail");
    expect(pane().textContent, "a sentence contradicting itself").not.toContain("still there");
    // And the wizard is usable again, which is the half that was terminal.
    expect(pane().querySelector(".step-of")?.textContent).toBe("Step 5 of 5 · Merge");
    expect(pane().querySelectorAll(".card.undo")).toHaveLength(0);
  });

  it("says so when the sweep left some of it behind", async () => {
    // A sweep that removed fewer slides than it asked for leaves part of the
    // preview in the deck, and the user is the only one who can finish it.
    await reachPreview();
    office.runMerge.mockResolvedValueOnce(PREVIEW);
    primary().click();
    await settle();

    office.undoMerge.mockResolvedValueOnce({ removed: 1, detail: "asked for 3 slide(s) and the deck shrank by 1" });
    // The deck agrees: two of the three are still in it.
    office.slideCount.mockResolvedValueOnce(14);
    primary().click();
    await settle();

    expect(said().join(" ")).toContain("still there");
    // And it is still a preview, so the merge stays blocked.
    expect(pane().querySelectorAll(".card.undo")).toHaveLength(1);
  });

  it("does not leave the button dead when the preview raises", async () => {
    await reachPreview();
    office.runMerge.mockRejectedValueOnce(new Error("the host would not export the slides"));
    primary().click();
    await settle();
    expect(said().join(" ")).toContain("would not export");
    expect(primary().disabled).toBe(false);
    expect(primary().textContent).toBe("Preview the first row");
  });

  it("cannot be started twice, and SAYS it is running", async () => {
    // Inserting a preview is a real merge and can take a minute on this host.
    // A button reading "Preview the first row", greyed out, for the whole of it
    // is the state a user cannot tell from a pane that has stopped responding —
    // and the other two long calls already named themselves.
    await reachPreview();
    const held = deferred<unknown>();
    office.runMerge.mockReturnValueOnce(held.promise);
    primary().click();
    await settle();
    expect(primary().textContent).toBe("Previewing…");
    expect(primary().disabled).toBe(true);
    primary().click();
    await settle();
    expect(office.runMerge).toHaveBeenCalledTimes(1);
    held.resolve(PREVIEW);
    await settle();
  });

  it("says it is REMOVING while the sweep is out", async () => {
    await reachPreview();
    office.runMerge.mockResolvedValueOnce(PREVIEW);
    primary().click();
    await settle();

    const held = deferred<unknown>();
    office.undoMerge.mockReturnValueOnce(held.promise);
    removePreview().click();
    await settle();
    expect(primary().textContent).toBe("Removing…");
    expect(primary().disabled).toBe(true);

    held.resolve({ removed: 3, detail: "removed 3 slide(s) from index 12" });
    await settle();
    expect(primary().textContent).toBe("Preview the first row");
  });
});

describe("filling the boxes from the slides the user selected", () => {
  it("puts the numbers in the boxes rather than committing a block", () => {
    // Two steps for one action would be worse, and skipping the template read
    // would be worse still — the fields step would then list nothing. What
    // this asserts is that the read is NOT spent: whether `state.block` is
    // also cleared is not observable and is not claimed.
    return (async () => {
      await openPane();
      await settle();
      office.selectedBlock.mockResolvedValueOnce({ ok: true, from: 4, to: 6 });
      (pane().querySelector('[data-action="selection"]') as HTMLElement).click();
      await settle();

      expect((field("from") as HTMLInputElement).value).toBe("4");
      expect((field("to") as HTMLInputElement).value).toBe("6");
      // Still on the template step, with the read not yet spent.
      expect(primary().textContent).toBe("Use slides 4 to 6");
      expect(office.inspectBlock).not.toHaveBeenCalled();
    })();
  });

  it("says what is wrong instead of filling anything", async () => {
    await openPane();
    await settle();
    office.selectedBlock.mockResolvedValueOnce({
      ok: false,
      why: "Slides 2 to 4 are not all selected — a template block has to be slides that sit next to each other.",
    });
    (pane().querySelector('[data-action="selection"]') as HTMLElement).click();
    await settle();

    expect(said().join(" ")).toContain("next to each other");
    expect((field("from") as HTMLInputElement).value).toBe("");
  });

  it("cannot be started twice, and gives the pane back either way", async () => {
    await openPane();
    await settle();
    const held = deferred<unknown>();
    office.selectedBlock.mockReturnValueOnce(held.promise);
    const link = () => pane().querySelector('[data-action="selection"]') as HTMLElement;
    link().click();
    await settle();
    expect(primary().disabled).toBe(true);
    link().click();
    await settle();
    expect(office.selectedBlock).toHaveBeenCalledTimes(1);

    held.resolve({ ok: true, from: 1, to: 2 });
    await settle();
    expect(primary().disabled).toBe(false);
  });

  it("clears a committed block, because the boxes now name different slides", async () => {
    await openPane();
    await settle();
    type("from", "1");
    type("to", "2");
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click();
    await settle();
    (pane().querySelector("[data-back]") as HTMLElement).click();

    office.selectedBlock.mockResolvedValueOnce({ ok: true, from: 7, to: 9 });
    (pane().querySelector('[data-action="selection"]') as HTMLElement).click();
    await settle();
    // The old block's placeholders must not stand behind the new numbers.
    expect(primary().textContent).toBe("Use slides 7 to 9");
  });
});

describe("a host too old to read the selection", () => {
  it("does not offer the shortcut at all", async () => {
    // `getSelectedSlides` is PowerPointApi 1.5 and the floor is 1.2, so this
    // is an EXTRA. It shipped unguarded — the call went in on a sibling
    // project's rounds showing it is not wedged, without anyone asking which
    // version introduced it. Safe to call and present are different questions.
    office.canReadSelection.mockReturnValue(false);
    await openPane();
    await settle();
    expect(pane().querySelector('[data-action="selection"]')).toBeNull();
    // And the boxes still work, which is the whole reason it can be absent.
    type("from", "4");
    type("to", "6");
    expect(primary().textContent).toBe("Use slides 4 to 6");
  });

  it("offers it where the host has it", async () => {
    await openPane();
    await settle();
    expect(pane().querySelector('[data-action="selection"]')).not.toBeNull();
  });
});

describe("taking rows out, through the real pane", () => {
  /** Reach the merge step with three rows pasted. */
  async function reachRows(): Promise<void> {
    await openPane();
    await settle();
    type("from", "4");
    type("to", "6");
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click();
    await settle();
    type("paste", "First\tLast\nAda\tLovelace\nGrace\tHopper\nKay\tMcNulty");
    primary().click(); // data -> fields
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // fields, re-reading the slides -> preview
    await settle();
    (pane().querySelector("[data-forward]") as HTMLElement).click(); // -> merge
    (pane().querySelector('[data-action="rows"]') as HTMLElement).click(); // open the list
    await settle();
  }

  const box = (i: number) => pane().querySelector(`[data-row="${i}"]`) as HTMLInputElement;

  it("takes a row out and puts the new number on the button", async () => {
    await reachRows();
    expect(primary().textContent).toBe("Add 9 slides");
    box(1).click();
    await settle();
    expect(primary().textContent).toBe("Add 6 slides");
    expect(box(1).checked).toBe(false);
  });

  it("puts it back on a second click", async () => {
    await reachRows();
    box(1).click();
    await settle();
    box(1).click();
    await settle();
    expect(primary().textContent).toBe("Add 9 slides");
  });

  it("merges only the rows left ticked", async () => {
    await reachRows();
    box(0).click();
    await settle();
    office.runMerge.mockResolvedValueOnce({ ...OUTCOME, added: 6 });
    primary().click();
    await settle();
    const req = office.runMerge.mock.calls[0]?.[0] as { records: { rows: { First: string }[] } };
    expect(req.records.rows.map((r) => r.First)).toEqual(["Grace", "Kay"]);
  });

  it("FORGETS the filter when a new table is pasted", async () => {
    // Row 2 of the old paste is not row 2 of the new one. Carrying an
    // exclusion across would take out a row the user never looked at.
    await reachRows();
    box(1).click();
    await settle();
    expect(primary().textContent).toBe("Add 6 slides");

    (pane().querySelector("[data-back]") as HTMLElement).click(); // preview
    (pane().querySelector("[data-back]") as HTMLElement).click(); // fields
    (pane().querySelector("[data-back]") as HTMLElement).click(); // data
    type("paste", "First\tLast\nZoe\tZed\nYan\tYates");
    primary().click(); // data -> fields
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // fields, re-reading the slides -> preview
    await settle();
    (pane().querySelector("[data-forward]") as HTMLElement).click();
    await settle();
    // Two rows, both in. The list stays OPEN across the paste — the disclosure
    // is a UI state about the pane, not about the data — so the boxes are
    // already on screen and clicking the toggle here would shut them.
    expect(primary().textContent).toBe("Add 6 slides");
    expect(box(0).checked).toBe(true);
    expect(box(1).checked).toBe(true);
    expect(pane().querySelectorAll(".rowlist li")).toHaveLength(2);
  });

  it("refuses a row click while a host call is out", async () => {
    // The answer on its way back is about the rows as they were when it left.
    // The BUTTON cannot show this — it reads "Merging…" either way — so the
    // checkbox is what says whether the click was taken.
    await reachRows();
    const held = deferred<unknown>();
    office.runMerge.mockReturnValueOnce(held.promise);
    primary().click();
    await settle();
    expect(primary().textContent).toBe("Merging…");

    box(1).click();
    await settle();
    expect(box(1).checked, "the click was refused").toBe(true);

    held.resolve({ ...OUTCOME, added: 9 });
    await settle();
    expect(box(1).checked).toBe(true);
  });
});

describe("taking a real merge back", () => {
  /** Walk to the merge step and land a run of six slides. */
  async function afterMerge(outcome: Record<string, unknown> = OUTCOME): Promise<HTMLElement> {
    const root = await openPane();
    await settle();
    type("from", "4");
    type("to", "6");
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click();
    await settle();
    type("paste", "First\tLast\nAda\tLovelace\nGrace\tHopper");
    primary().click(); // data -> fields
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // fields, re-reading the slides -> preview
    await settle();
    // The preview step's primary RUNS a preview; the way past it is its own
    // control. Clicking the primary here ran a merge and left the pane on
    // the preview step, which is why the first version of this helper landed
    // nothing.
    document.querySelector<HTMLButtonElement>('[data-forward="merge"]')?.click();
    office.runMerge.mockResolvedValueOnce(outcome);
    primary().click();
    await settle();
    return root;
  }

  const undoButton = (): HTMLButtonElement | null =>
    document.querySelector<HTMLButtonElement>('.card.undo button[data-action="undo"]');

  /**
   * Put the merge button back within reach.
   *
   * A landed merge disarms it — pressing it twice by accident is the thing that
   * guard exists to stop — so a second run needs the user to change something
   * first. Taking a row out is the cheapest change that does it, and it is what
   * somebody who has just looked at the output would actually do.
   */
  async function rearm(): Promise<void> {
    (pane().querySelector('[data-action="rows"]') as HTMLElement).click();
    await settle();
    (pane().querySelector('[data-row="1"]') as HTMLInputElement).click();
    await settle();
  }

  it("offers the way back once a merge has landed", async () => {
    // `undoInsert` and `sweepPlan` were built and tested before this and were
    // reachable from nothing — the numbers were kept, the sentence was
    // written, and no view rendered either.
    await afterMerge();
    expect(undoButton()).not.toBeNull();
  });

  it("does NOT offer it when the run cannot say which slides are its own", async () => {
    // The deck grew by more than the package held — a co-author or AutoSave
    // landing a slide across the insert. The slides are there and some are
    // ours; nothing here can say which, and `sweepPlan` refuses that shape
    // deliberately. Offering the card anyway put a slide-deleting button on
    // screen that answered "nothing to take back" every time it was pressed.
    await afterMerge({
      ...OUTCOME,
      ok: false,
      accountable: false,
      detail: "The deck changed in a way this run cannot account for.",
    });
    expect(undoButton(), "an offer the sweep will decline").toBeNull();
  });

  it("sweeps with the run's OWN numbers, not the pane's current ones", async () => {
    // The clamps are only worth anything against the count taken before the
    // run inserted. Handing the sweep a number the pane happens to hold now is
    // how a positional delete reaches slides the user owned first.
    await afterMerge();
    office.undoMerge.mockResolvedValueOnce({ removed: 6, detail: "removed 6 slide(s) from index 12" });
    // The deck's size is ASKED for after a sweep rather than computed from the
    // pane's cached number, which is stale exactly when the user has been
    // editing by hand.
    office.slideCount.mockResolvedValueOnce(12);
    undoButton()?.click();
    await settle();

    expect(office.undoMerge).toHaveBeenCalledTimes(1);
    expect(office.undoMerge.mock.calls[0]?.[0]).toMatchObject({ deckAtStart: 12, added: 6 });
  });

  it("puts the way back away once the slides are gone", async () => {
    await afterMerge();
    office.undoMerge.mockResolvedValueOnce({ removed: 6, detail: "removed 6 slide(s) from index 12" });
    // The deck's size is ASKED for after a sweep rather than computed from the
    // pane's cached number, which is stale exactly when the user has been
    // editing by hand.
    office.slideCount.mockResolvedValueOnce(12);
    undoButton()?.click();
    await settle();
    expect(undoButton()).toBeNull();
    // The deck AFTER, measured from what came out. "Back to 12" is only true
    // when the sweep took everything, and it said so over a deck of 14.
    expect(document.body.textContent).toContain("Your deck holds 12");
  });

  it("does not let a second merge that added NOTHING destroy the way back", async () => {
    /**
     * The card and the button read from two different places — the sentence
     * from `state.added`, the sweep from the module's `last` — and only one of
     * them was guarded. A run that added nothing overwrote `last` and cleared
     * the crumb, while `state.added` correctly kept the first run's numbers.
     *
     * So the card went on offering "Remove slides 13 to 18", the button stayed
     * live, and pressing it swept with the second run's numbers and removed
     * nothing, forever. Six slides in the deck, no way back to them, and no
     * crumb either if the tab then died. Two sources of truth for one offer.
     */
    await afterMerge();
    expect(undoButton(), "the first run is offered").not.toBeNull();

    // The user changes their mind about a row, which re-arms the button, and
    // merges again. The host refuses and nothing lands.
    await rearm();
    office.runMerge.mockResolvedValueOnce({ ...OUTCOME, ok: false, added: 0, deckAtStart: 18, runId: "r2" });
    primary().click();
    await settle();
    expect(office.runMerge, "the second merge really ran").toHaveBeenCalledTimes(2);

    office.undoMerge.mockResolvedValueOnce({ removed: 6, detail: "removed 6 slide(s) from index 12" });
    // The deck's size is ASKED for after a sweep rather than computed from the
    // pane's cached number, which is stale exactly when the user has been
    // editing by hand.
    office.slideCount.mockResolvedValueOnce(12);
    undoButton()?.click();
    await settle();
    expect(office.undoMerge.mock.calls[0]?.[0], "the six slides that are actually there").toMatchObject({
      deckAtStart: 12,
      added: 6,
    });
  });

  it("keeps the crumb for slides that are still in the deck", async () => {
    // Same defect, seen from the record a dead tab leaves behind. The pending
    // marker written at the start of every run overwrote a crumb describing
    // six slides nobody had taken back yet.
    await afterMerge();
    await rearm();
    office.runMerge.mockResolvedValueOnce({ ...OUTCOME, ok: false, added: 0, deckAtStart: 18, runId: "r2" });
    primary().click();
    await settle();

    const stored: unknown = JSON.parse(globalThis.localStorage.getItem(CRUMB_KEY) ?? "null");
    expect(stored, "the first run's six slides are still recorded").toMatchObject({ deckAtStart: 12, added: 6 });
  });

  it("will not offer to remove more slides than the merge could possibly have added", async () => {
    /**
     * `runMerge` caps `added` at the size of the package it sent, and the
     * comment at that line says why: an uncapped count absorbs whatever else
     * arrived, so `grew` and `added` are equal by construction and the clamp
     * keeping an undo off a stranger's slides can never fire.
     *
     * The RAISE path recomputed the same quantity from the deck with no cap at
     * all — and it is the likelier path for this, because a co-author's slides
     * arriving is exactly the kind of thing that also makes a call time out.
     * Deck of 12, a merge of six slides that raises, twelve slides from
     * somebody else landing meanwhile: the pane counted 30, called it 18, and
     * offered to remove eighteen.
     */
    const root = await openPane();
    await settle();
    type("from", "4");
    type("to", "6");
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click();
    await settle();
    type("paste", "First\tLast\nAda\tLovelace\nGrace\tHopper");
    primary().click();
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click();
    await settle();
    document.querySelector<HTMLButtonElement>('[data-forward="merge"]')?.click();
    expect(primary().textContent, "two rows over a three-slide block").toBe("Add 6 slides");

    // The insert raises. The deck is 30 by the time the pane counts again:
    // this run's six, and twelve from somebody else.
    office.slideCount.mockResolvedValueOnce(12).mockResolvedValueOnce(30);
    office.runMerge.mockRejectedValueOnce(new Error("gave up waiting for: inserting the merged deck"));
    primary().click();
    await settle();
    expect(root.textContent, "the raise is reported").toContain("landed anyway");

    // Whatever it offers, it may never claim more than the merge could build.
    office.undoMerge.mockResolvedValueOnce({ removed: 0, detail: "nothing to take back" });
    undoButton()?.click();
    await settle();
    const asked = office.undoMerge.mock.calls[0]?.[0] as { added: number } | undefined;
    expect(asked?.added ?? 0, "never more than the six slides this merge builds").toBeLessThanOrEqual(6);
  });

  it("does not leave a recovered run's card over a LATER merge's steps", async () => {
    /**
     * `recovered` says the offer follows the SLIDES rather than the merge step,
     * because a run whose pane died leaves the user wherever they happen to be.
     * It was set at boot and never cleared, so it stayed true for the rest of
     * the session — including for an ORDINARY merge started afterwards, whose
     * card then drew on every step of the wizard.
     *
     * A button that deletes slides out of a presentation belongs on the step
     * that made them. On the template step it is a control with no context at
     * all, one press from removing part of the deck the user is choosing from.
     */
    localStorage.setItem(
      CRUMB_KEY,
      JSON.stringify({
        kind: "ssf-merge-run",
        deckAtStart: 12,
        added: 6,
        runId: "died",
        startedAt: "2026-08-27T10:00:00.000Z",
        doc: "https://example-my.sharepoint.com/personal/x/Documents/deck.pptx",
      }),
    );
    office.slideCount.mockReset().mockResolvedValue(18);
    await openPane();
    await settle();
    expect(undoButton(), "the dead run is offered, wherever the user is").not.toBeNull();

    // The user takes those slides back, and then does an ordinary merge.
    office.undoMerge.mockResolvedValueOnce({ removed: 6, detail: "removed 6 slide(s) from index 12" });
    // The deck's size is ASKED for after a sweep rather than computed from the
    // pane's cached number, which is stale exactly when the user has been
    // editing by hand.
    office.slideCount.mockResolvedValueOnce(12);
    undoButton()?.click();
    await settle();
    office.slideCount.mockReset().mockResolvedValue(12);

    // An ordinary merge, in the SAME pane. Re-opening would reset the module's
    // state and with it the flag under test, so the walk is done by hand.
    type("from", "4");
    type("to", "6");
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click();
    await settle();
    type("paste", "First\tLast\nAda\tLovelace\nGrace\tHopper");
    primary().click();
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click();
    await settle();
    document.querySelector<HTMLButtonElement>('[data-forward="merge"]')?.click();
    office.runMerge.mockResolvedValueOnce(OUTCOME);
    primary().click();
    await settle();
    expect(undoButton(), "the new merge's own way back").not.toBeNull();

    for (let i = 0; i < STEP_COUNT; i++) {
      pane().querySelector<HTMLElement>("[data-back]")?.click();
      await settle();
    }
    expect(document.body.textContent, "walked all the way back").toContain("Step 1 of 5");
    expect(undoButton(), "not on the template step").toBeNull();
  });

  it("asks the deck for its size rather than adjusting the number it had", () => {
    /**
     * The pane's cached size is stale exactly when the user has edited the deck
     * by hand — which is the only way to reach the sentence that prints it. It
     * said "your deck holds 15" over a deck of 14 and then wrote 15 into the
     * state, so the merge card went on to offer "6 slides added after slide 15,
     * leaving 21" over a fourteen-slide deck.
     *
     * The two sibling paths in this file already re-count. This asserts the
     * third does, by giving the host an answer nothing else could produce.
     */
    return (async () => {
      await afterMerge();
      // A COMPLETE sweep, which is the branch that prints the size.
      office.undoMerge.mockResolvedValueOnce({ removed: 6, disowned: 0, detail: "removed 6" });
      office.slideCount.mockResolvedValueOnce(99);
      undoButton()?.click();
      await settle();
      // 99 is not 12 + 6 - 6, so only a real read can produce it.
      expect(pane().textContent).toContain("Your deck holds 99");
    })();
  });

  it("KEEPS the way back when the sweep only got some of them", async () => {
    // A partial sweep leaves slides in the deck and the user is the only one
    // who can finish the job, so the button has to stay.
    await afterMerge();
    office.undoMerge.mockResolvedValueOnce({
      removed: 2,
      disowned: 0,
      detail: "asked for 6 and the deck shrank by 2",
    });
    // The deck is ASKED for its size after a sweep, rather than the pane's
    // cached number being adjusted — the cache is stale exactly when the user
    // has edited the deck by hand. Six added onto twelve, two taken back.
    office.slideCount.mockResolvedValueOnce(16);
    undoButton()?.click();
    await settle();
    expect(undoButton(), "still offered").not.toBeNull();
    expect(document.body.textContent).toContain("Some of the merge is still there");
  });

  it("does not go on offering slides the sweep has DISOWNED", async () => {
    /**
     * `provenSweep` leaves a slide in the range that carries no mark of this
     * run — the user deleted two merged slides and appended two of their own.
     * `added - removed` counted those as still owed, so the card stayed up
     * saying "Remove slides 13 to 14, which this merge added" over slides the
     * very same sentence had just called not this merge's, with a live delete
     * button on them.
     */
    await afterMerge();
    office.undoMerge.mockResolvedValueOnce({
      removed: 4,
      disowned: 2,
      detail:
        "removed 4 slide(s) from slides 13 to 18; 2 slide(s) in the range are not this merge's and were left alone",
    });
    undoButton()?.click();
    await settle();
    expect(undoButton(), "an offer over slides the sweep disowned").toBeNull();
    expect(document.body.textContent, "and it does not claim they are still owed").not.toContain(
      "Some of the merge is still there",
    );
    // What was DECLINED, counted. "The rest could not be shown to be this
    // merge's" spoke about every slide left in the range — four of them here,
    // three of which the sweep never doubted — so the sentence written to be
    // true on every path was false on this one.
    expect(document.body.textContent).toContain("2 slides in that range could not be shown to be this merge's");
    expect(document.body.textContent, "a claim about slides the sweep never doubted").not.toContain("The rest");
  });

  it("does not offer a run again once a press has proved it cannot be taken back", async () => {
    /**
     * The crumb is kept after a withdrawal — it is the record that stops the
     * next merge overwriting slides that are still in the deck — and it used to
     * be kept VERBATIM. So every future open of that deck said "the pane closed
     * before you could take them back", about a press that had happened and
     * been refused, over a card that died the moment it was pressed. Once per
     * open, indefinitely, under a sentence that was not true.
     */
    localStorage.setItem(
      CRUMB_KEY,
      JSON.stringify({
        kind: "ssf-merge-run",
        deckAtStart: 12,
        added: 6,
        runId: "r1",
        startedAt: "2026-08-27T10:00:00.000Z",
        doc: "https://example-my.sharepoint.com/personal/x/Documents/deck.pptx",
        pressed: true,
        unremovable: true,
      }),
    );
    office.slideCount.mockReset().mockResolvedValue(18);
    await openPane();
    await settle();

    expect(undoButton(), "a card that would die on its first press").toBeNull();
    expect(document.body.textContent, "a sentence that is not true").not.toContain("the pane closed");
    expect(document.body.textContent).toContain("could not take back");
    expect(document.body.textContent).toContain("thumbnail rail");
  });

  it("does not say the pane closed when a press has already been made and answered", async () => {
    /**
     * Found by the real-host round of 2026-09-02, on PowerPoint for the web.
     *
     * The user merged six slides, deleted all six from the thumbnail rail, and
     * pressed "Remove these slides". The pane answered, correctly, "Nothing was
     * removed - nothing to take back (deck was 13, is 13)". The crumb was kept
     * and marked `pressed`, which is right: it is the record that stops the
     * next merge overwriting a run whose slides may still be there.
     *
     * On the next open the pane said "A merge from <date> added 6 slides and
     * the pane closed before you could take them back." Two things wrong with
     * one sentence. The pane did not close before the press - the press
     * happened and was answered. And the slides are not outstanding: the deck
     * is back to its starting size, so no card is offered beside the sentence
     * that says they are waiting to be taken back.
     *
     * `unremovable` was already handled. `pressed` was not, and it is the
     * commoner mark by far: it is written on EVERY fruitless press, while
     * `unremovable` needs the budget spent or a host with no tags at all.
     */
    localStorage.setItem(
      CRUMB_KEY,
      JSON.stringify({
        kind: "ssf-merge-run",
        deckAtStart: 12,
        added: 6,
        runId: "r1",
        startedAt: "2026-09-02T10:00:00.000Z",
        doc: "https://example-my.sharepoint.com/personal/x/Documents/deck.pptx",
        pressed: true,
        fruitless: 1,
      }),
    );
    // The deck is back to the size it was before the merge, so `sweepPlan`
    // refuses and no card is drawn. The sentence must not contradict that.
    office.slideCount.mockReset().mockResolvedValue(12);
    await openPane();
    await settle();

    expect(undoButton(), "the deck did not grow, so there is nothing to offer").toBeNull();
    expect(document.body.textContent, "the press happened; the pane did not close first").not.toContain(
      "the pane closed",
    );
  });

  it("does not report the sweep's remainder as what the merge added", async () => {
    // `added` is what a further press may still take back, so a partial undo
    // lowers it — and the disabled merge button read from it, so a six-slide
    // merge with three swept back said "Added 3 slides" about a merge that
    // added six. Two facts, one number.
    await afterMerge();
    expect(primary().textContent).toBe("Added 6 slides");

    office.undoMerge.mockResolvedValueOnce({ removed: 3, disowned: 0, detail: "removed 3 slide(s)" });
    office.slideCount.mockResolvedValueOnce(15);
    undoButton()?.click();
    await settle();

    expect(primary().textContent, "a merge that added six").not.toBe("Added 3 slides");
    expect(primary().textContent).toBe("3 of 6 slides still there");
    expect(primary().disabled, "and it is still disarmed").toBe(true);
  });

  it("marks a press that moved nothing as a press, so the next one still asks for proof", async () => {
    /**
     * `pressed` was set only where slides came out. A press that removed
     * nothing left the next one looking like a FIRST press — no proof asked —
     * and `provenSweep`'s pre-tags fall-through then takes the whole positional
     * window, which is where a slide the user made in between sits.
     *
     * Reachable without any host misbehaviour worth the name: PowerPoint
     * accepts the deletes and performs none, which `undoInsert` already guards
     * for, and answers `removed: 0`.
     */
    await afterMerge();
    office.undoMerge.mockResolvedValueOnce({
      removed: 0,
      disowned: 2,
      detail: "asked for 4 slide(s) from slides 13 to 18 and the deck shrank by 0",
    });
    office.slideCount.mockResolvedValueOnce(18);
    undoButton()?.click();
    await settle();
    expect(undoButton(), "one failed press is not a host that cannot answer").not.toBeNull();

    office.undoMerge.mockResolvedValueOnce({ removed: 6, disowned: 0, detail: "removed 6 slide(s)" });
    office.slideCount.mockResolvedValueOnce(12);
    undoButton()?.click();
    await settle();
    expect(office.undoMerge.mock.calls[1]?.[0], "the second press is not a first one").toMatchObject({
      pressed: true,
    });
  });

  it("stops offering after two presses that prove nothing, on a host that could have proved", async () => {
    /**
     * The other half of the same trade. A host that HAS slide tags and does not
     * answer with them looks exactly like one that failed a single read — so
     * the first fruitless press keeps the offer, and an unbounded number of
     * them would leave a delete button standing over slides no press can take.
     * Two is the smallest number that tells a hiccup from a state.
     */
    await afterMerge();
    const fruitless = {
      removed: 0,
      disowned: 6,
      detail: "nothing to take back — none of slides 13 to 18 could be shown to be this merge's",
    };
    office.undoMerge.mockResolvedValueOnce(fruitless);
    office.slideCount.mockResolvedValueOnce(18);
    undoButton()?.click();
    await settle();
    expect(undoButton(), "the first one may have been a bad minute").not.toBeNull();

    office.undoMerge.mockResolvedValueOnce(fruitless);
    office.slideCount.mockResolvedValueOnce(18);
    undoButton()?.click();
    await settle();
    expect(undoButton(), "twice is a state, not a minute").toBeNull();
    expect(document.body.textContent).toContain("thumbnail rail");
  });

  it("counts a press the host swallowed, not only one that disowned something", async () => {
    /**
     * The other case `FRUITLESS_LIMIT` exists for. PowerPoint accepts the
     * deletes and performs none: `undoInsert` proved the whole plan, so nothing
     * is disowned and nothing came out — `removed: 0, disowned: 0`. Counting
     * only the disowned shape meant this one never advanced the counter, and
     * the offer stood over slides no press could take for the rest of the
     * session.
     */
    await afterMerge();
    const swallowed = {
      removed: 0,
      disowned: 0,
      detail: "asked for 6 slide(s) from slides 13 to 18 and the deck shrank by 0",
    };
    office.undoMerge.mockResolvedValueOnce(swallowed);
    office.slideCount.mockResolvedValueOnce(18);
    undoButton()?.click();
    await settle();
    expect(undoButton(), "once may be a bad minute").not.toBeNull();

    office.undoMerge.mockResolvedValueOnce(swallowed);
    office.slideCount.mockResolvedValueOnce(18);
    undoButton()?.click();
    await settle();
    expect(undoButton(), "twice is a host that will not do it").toBeNull();
  });

  it("does not spend the budget on a deck the sweep refused the shape of", async () => {
    /**
     * `sweepPlan` declines before anything is asked of PowerPoint — the deck
     * grew past what this run added, which is a co-author's slide, not a host
     * misbehaving. Counting it left one genuine hiccup enough to write the
     * record off permanently.
     */
    await afterMerge();
    office.undoMerge.mockResolvedValueOnce({
      removed: 0,
      disowned: 0,
      refusedShape: true,
      detail: "nothing to take back (deck was 12, is 30)",
    });
    office.slideCount.mockResolvedValueOnce(18);
    undoButton()?.click();
    await settle();

    // The budget is untouched, so two real fruitless presses are still needed.
    const swallowed = { removed: 0, disowned: 0, detail: "the deck shrank by 0" };
    office.undoMerge.mockResolvedValueOnce(swallowed);
    office.slideCount.mockResolvedValueOnce(18);
    undoButton()?.click();
    await settle();
    expect(undoButton(), "one hiccup is not a state").not.toBeNull();
  });

  it("gives the budget back after a press that worked", async () => {
    /**
     * `fruitless` was spread through the success path with the rest of the
     * outcome, so a fruitless press, a working one, and a second fruitless one
     * withdrew the card — a press short of the budget the changelog and the
     * manual both promise, with three merged slides still in the deck.
     */
    await afterMerge();
    office.undoMerge.mockResolvedValueOnce({ removed: 0, disowned: 6, detail: "nothing to take back" });
    office.slideCount.mockResolvedValueOnce(18);
    undoButton()?.click();
    await settle();

    office.undoMerge.mockResolvedValueOnce({ removed: 3, disowned: 0, detail: "removed 3 slide(s)" });
    office.slideCount.mockResolvedValueOnce(15);
    undoButton()?.click();
    await settle();
    expect(undoButton(), "slides are still owed").not.toBeNull();

    office.undoMerge.mockResolvedValueOnce({ removed: 0, disowned: 3, detail: "nothing to take back" });
    office.slideCount.mockResolvedValueOnce(15);
    undoButton()?.click();
    await settle();
    expect(undoButton(), "the count was not given back by the press that worked").not.toBeNull();
  });

  it("does not carry a withdrawal into the next merge, on either path out of a run", async () => {
    /**
     * `undoWithdrawn` is about the merge that earned it. The success path
     * cleared it; the path where the merge RAISES and the slides land anyway
     * did not — so a withdrawal, an edit, and a second merge that raised left
     * nine slides in the deck with `last` correctly set and no card drawn. The
     * only way back was gone for the session.
     */
    await afterMerge();
    office.undoMerge.mockResolvedValueOnce({
      removed: 0,
      disowned: 6,
      unprovable: true,
      detail: "nothing to take back — none of slides 13 to 18 could be shown to be this merge's",
    });
    office.slideCount.mockResolvedValueOnce(18);
    undoButton()?.click();
    await settle();
    expect(undoButton(), "withdrawn").toBeNull();

    // An edit is a different merge, so the button arms again.
    await rearm();
    office.slideCount.mockResolvedValueOnce(18).mockResolvedValueOnce(21); // before, after
    office.runMerge.mockRejectedValueOnce(new Error("gave up waiting for: inserting the merged deck"));
    primary().click();
    await settle();

    expect(document.body.textContent, "the slides landed").toContain("landed anyway");
    expect(undoButton(), "a way back to the slides this merge left").not.toBeNull();
  });

  it("takes the card down only where the host says the press can NEVER work", async () => {
    /**
     * `removed: 0, disowned: n` is not a terminal answer. A 1.3 host that
     * failed one tag read, and a delete PowerPoint accepted and did not
     * perform, both produce it and both succeed on the next press — so a
     * withdrawal on that pair alone throws away slides the very next press
     * would have removed. `undoInsert` says which case this is: `unprovable`
     * is set only where proof was required and the host has no slide tags at
     * all.
     */
    await afterMerge();
    office.undoMerge.mockResolvedValueOnce({ removed: 2, disowned: 0, detail: "removed 2 slide(s)" });
    office.slideCount.mockResolvedValueOnce(16);
    undoButton()?.click();
    await settle();
    expect(undoButton(), "slides are still owed, so the way back stays").not.toBeNull();

    // A press that proved nothing but might next time. The card stays.
    office.undoMerge.mockResolvedValueOnce({
      removed: 0,
      disowned: 4,
      detail: "nothing to take back — none of slides 13 to 16 could be shown to be this merge's",
    });
    office.slideCount.mockResolvedValueOnce(16);
    undoButton()?.click();
    await settle();
    expect(undoButton(), "a host that answered nothing once may answer the next time").not.toBeNull();

    // The same answer from a host that CANNOT answer. Now it goes.
    office.undoMerge.mockResolvedValueOnce({
      removed: 0,
      disowned: 4,
      unprovable: true,
      detail: "nothing to take back — none of slides 13 to 16 could be shown to be this merge's",
    });
    office.slideCount.mockResolvedValueOnce(16);
    undoButton()?.click();
    await settle();

    expect(undoButton(), "a button that can only answer the same way again").toBeNull();
    // The slides are still in the deck, so the user is told what they can do
    // about them instead of being left with a sentence and no way forward.
    expect(document.body.textContent).toContain("thumbnail rail");
  });

  it("leaves the merge button disarmed when it takes the undo card down", async () => {
    /**
     * The card was hidden by clearing `state.added`, and `added` is what
     * disarms the merge button — its own docstring says so: "one more press and
     * there are 1440, in somebody's deck, from a button that looks like the one
     * they just pressed". So the withdrawal re-armed "Add 6 slides" over the
     * six slides that were still there, and invited the user to double them.
     *
     * The card and the button are different questions: one is what may be
     * pressed, the other is what is in the deck.
     */
    await afterMerge();
    expect(primary().textContent, "a merge that landed").toBe("Added 6 slides");

    office.undoMerge.mockResolvedValueOnce({
      removed: 0,
      disowned: 6,
      unprovable: true,
      detail: "nothing to take back — none of slides 13 to 18 could be shown to be this merge's",
    });
    office.slideCount.mockResolvedValueOnce(18);
    const merges = office.runMerge.mock.calls.length;
    undoButton()?.click();
    await settle();

    expect(undoButton(), "the card is withdrawn").toBeNull();
    expect(primary().textContent, "and the merge button is still the one that landed").toBe("Added 6 slides");
    expect(primary().disabled, "a live merge button over slides that are still there").toBe(true);
    primary().click();
    await settle();
    expect(office.runMerge.mock.calls.length, "no second merge on top of the first").toBe(merges);
  });

  it("says so when the sweep refused, and leaves the deck alone", async () => {
    // `sweepPlan` refuses when the deck gained more than the run added,
    // because the last N slides are then somebody else's. That refusal must
    // reach the user as a sentence rather than as a silent no-op.
    await afterMerge();
    office.undoMerge.mockResolvedValueOnce({
      removed: 0,
      disowned: 0,
      detail: "nothing to take back (deck was 12, is 20)",
    });
    // The deck says the same thing the detail does, and the pane asks it. It
    // used to keep the size it had at merge time, so the refusal printed
    // "deck was 12, is 20" beside a live "Remove slides 13 to 18" drawn from
    // an arithmetic that still believed 18 — two deck sizes on one screen, and
    // a destructive button that would refuse for ever.
    office.slideCount.mockResolvedValueOnce(20);
    undoButton()?.click();
    await settle();
    expect(document.body.textContent).toContain("Nothing was removed");
    expect(undoButton(), "an offer that can only refuse again").toBeNull();
  });

  it("KEEPS the way back when the user edits the merge after it landed", async () => {
    // `added` did two jobs: it disarmed the button, and it was the only thing
    // the undo card was drawn from. Every edit cleared it — deliberately, so a
    // different merge can be pressed — and took the way back with it. The
    // slides stayed in the deck; the only offer to remove them left the screen
    // on a keystroke, and `main.ts` still held the numbers to do it.
    await afterMerge();
    expect(undoButton()).not.toBeNull();

    // The row list is on this very screen, so the user never leaves the merge
    // step to lose the offer.
    document.querySelector<HTMLButtonElement>('[data-action="rows"]')?.click();
    const box = document.querySelector<HTMLInputElement>('input[data-row="0"]');
    box?.click();
    await settle();

    expect(primary().disabled, "a different merge is pressable again").toBe(false);
    expect(undoButton(), "and the slides that landed are still offered back").not.toBeNull();
  });

  it("offers the way back when the merge RAISED and slides landed anyway", async () => {
    // The catch branch kept `added` and dropped `deckAtStart`, three lines
    // under a comment saying the two travel together everywhere because the
    // card needs both. So the one case the branch exists for — a host that
    // performs a call and then raises on it — was the one case with no offer.
    await reachMerge();
    office.slideCount.mockResolvedValueOnce(12).mockResolvedValueOnce(18); // before, after
    office.runMerge.mockRejectedValueOnce(new Error("gave up waiting for: inserting the merged deck"));
    primary().click();
    await settle();

    expect(document.body.textContent).toContain("landed anyway");
    expect(undoButton(), "six slides in the deck and a way to remove them").not.toBeNull();
    office.undoMerge.mockResolvedValueOnce({ removed: 6, detail: "removed 6 slide(s) from index 12" });
    // The deck's size is ASKED for after a sweep rather than computed from the
    // pane's cached number, which is stale exactly when the user has been
    // editing by hand.
    office.slideCount.mockResolvedValueOnce(12);
    undoButton()?.click();
    await settle();
    expect(office.undoMerge.mock.calls[0]?.[0]).toMatchObject({ deckAtStart: 12, added: 6 });
  });

  it("clamps the recovery to the deck as it was, not as the pane remembered it", async () => {
    /**
     * The pane counts the deck when the block is committed, and the step
     * between that and this button is the one that sends the user into
     * PowerPoint to put fields on the slides. Two slides added there leave the
     * cached number low — and every clamp in `sweepPlan` compares SIZES, none
     * of them freshness, so a floor two slides behind yields a plan that starts
     * inside the user's own slides and satisfies every guard on the way.
     *
     * Cached 12, really 14, deck 20 after the raise. The run added six. Read
     * from the cache it would report eight and offer to delete from index 12.
     */
    await reachMerge();
    office.slideCount.mockResolvedValueOnce(14).mockResolvedValueOnce(20);
    office.runMerge.mockRejectedValueOnce(new Error("gave up waiting for: inserting the merged deck"));
    primary().click();
    await settle();

    expect(said().join(" "), "counted from the cache").toContain("6 slides landed anyway");
    office.undoMerge.mockResolvedValueOnce({ removed: 6, detail: "removed 6 slide(s) from index 14" });
    undoButton()?.click();
    await settle();
    expect(office.undoMerge.mock.calls[0]?.[0]).toMatchObject({ deckAtStart: 14, added: 6 });
  });

  it("does not take a second press while a sweep is out", async () => {
    await afterMerge();
    const held = deferred<unknown>();
    office.undoMerge.mockReturnValueOnce(held.promise);
    undoButton()?.click();
    await settle();
    undoButton()?.click();
    await settle();
    expect(office.undoMerge).toHaveBeenCalledTimes(1);
    held.resolve({ removed: 6, detail: "removed 6" });
    await settle();
  });
});

describe("the build stamp", () => {
  /**
   * Which build the host actually served, readable BEFORE anything is run.
   *
   * It already reaches the run record through `hostEnvironment()`, and that was
   * treated as enough. It is not: a run record exists only once a run has
   * finished, and the question is asked before one starts. PowerPoint caches
   * the pane's HTML for about ten minutes, so opening it too soon after a
   * deploy tests code the host never fetched and reads as a clean run of the
   * wrong build.
   *
   * The case that made it worth wiring: the build before it never loaded
   * Office.js and rendered a header and nothing else. Anyone testing that fix
   * has the broken build cached, and the two are told apart by this line or not
   * at all.
   */
  it("is in the header as soon as the pane opens, with no run needed", async () => {
    await openPane();
    const stamp = document.querySelector("header .build");
    expect(stamp?.textContent).toBe("test");
    // Before the floor check, so a host that REFUSES to run still says which
    // build refused. Asserted through the pane being unable to have run: no
    // merge has happened, so there is no run record for it to have come from.
    expect(office.runMerge).not.toHaveBeenCalled();
  });

  it("says nothing rather than 'unknown' when the build is not stamped", async () => {
    // A checkout with no git builds with `__BUILD_STAMP__` undefined and
    // `environmentLine` answers "unknown". A header reading `unknown` is worse
    // than an empty one: it looks like an answer.
    office.hostEnvironment.mockReturnValueOnce({
      build: "unknown",
      platform: "PC",
      host: "16.0.0",
      sets: ["1.2"],
      floor: "1.2",
      clearsFloor: true,
      deckSource: "file" as const,
      canSelect: true,
    });
    await openPane();
    expect(document.querySelector("header .build")).toBeNull();
  });
});

describe("the conditional slide control", () => {
  /**
   * The gap this closes: `prepare.ts` implemented conditional slides,
   * `PaneState` carried `conditions`, `main.ts` passed it to both the preview
   * and the merge — and nothing WROTE it, so the field was undefined in every
   * run that had ever happened. Built, tested, and reachable from nothing, for
   * the third time in this repo.
   *
   * So these tests drive the real control through the real `main.ts`: what a
   * user does is choose a column from a dropdown, and the thing that must be
   * true is that the choice reaches `runMerge`.
   */
  async function toMergeWithData(): Promise<void> {
    await reachMerge();
  }

  /** Walk back to the slide-number boxes, whichever step the pane is on. */
  function backToTemplate(): void {
    for (let i = 0; i < STEP_COUNT; i++) {
      if (pane().querySelector('[data-field="from"]')) return;
      (pane().querySelector("[data-back]") as HTMLElement).click();
    }
  }

  /** Walk back to the paste box. */
  function backToData(): void {
    for (let i = 0; i < STEP_COUNT; i++) {
      if (pane().querySelector('[data-field="paste"]')) return;
      (pane().querySelector("[data-back]") as HTMLElement).click();
    }
  }

  function openConditions(): void {
    (pane().querySelector('[data-action="conditions"]') as HTMLElement).click();
  }

  function choose(slide: number, column: string): void {
    const node = pane().querySelector(`[data-condition="${slide}"]`) as HTMLSelectElement;
    node.value = column;
    node.dispatchEvent(new Event("input", { bubbles: true }));
  }

  it("is not offered on the steps before the merge", async () => {
    // It lives beside the row picker, because "which rows" and "which slides"
    // are the two questions the merge screen is for. On the fields step — where
    // it used to be — it sat next to a control for putting placeholders onto
    // slides, which is a different job.
    await openPane();
    await settle();
    type("from", "4");
    type("to", "6");
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click();
    await settle();
    expect(pane().querySelector('[data-action="conditions"]'), "on the data step").toBeNull();
    type("paste", "First\tLast\nAda\tLovelace\nGrace\tHopper");
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // data -> fields
    expect(pane().querySelector('[data-action="conditions"]'), "on the fields step").toBeNull();
  });

  it("offers one control per slide in the block, defaulting to always", async () => {
    await toMergeWithData();
    openConditions();
    const selects = Array.from(pane().querySelectorAll("[data-condition]"));
    expect(selects.map((s) => s.getAttribute("data-condition"))).toEqual(["4", "5", "6"]);
    expect(selects.every((s) => (s as HTMLSelectElement).value === "")).toBe(true);
    // The columns from THIS paste, and nothing invented.
    const options = Array.from(selects[0]!.querySelectorAll("option")).map((o) => o.value);
    expect(options).toEqual(["", "First", "Last"]);
  });

  it("carries the choice into the merge", async () => {
    // The assertion the whole feature turns on. Everything else here could pass
    // while `runMerge` was still handed nothing.
    await toMergeWithData();
    openConditions();
    choose(5, "Last");
    office.runMerge.mockResolvedValueOnce(OUTCOME);
    primary().click();
    await settle();
    expect(office.runMerge).toHaveBeenCalledTimes(1);
    expect((office.runMerge.mock.calls[0]?.[0] as { conditions?: unknown }).conditions).toEqual({ 5: "Last" });
  });

  it("forgets conditions when the block moves", async () => {
    /**
     * Keyed by SLIDE NUMBER, so "slide 5 only when Last" is about the fifth
     * slide of the deck. Carried across a changed block it lands on whichever
     * slide now holds that number — a slide the user never set a condition on,
     * silently, discovered by counting slides in the output.
     *
     * The block here OVERLAPS the old one (4-6 becomes 3-5) and that is the
     * whole point of the case. A first version moved to 7-9, where the stale
     * key 5 is outside the new block and the engine ignores it — so the test
     * passed with the clearing removed and proved nothing. Check which
     * assertion goes red, and against what.
     */
    await toMergeWithData();
    openConditions();
    choose(5, "Last");
    backToTemplate();
    type("from", "3");
    type("to", "5");
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // template -> data
    await settle();
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // data -> fields
    primary().click(); // fields -> preview
    await settle();
    (pane().querySelector("[data-forward]") as HTMLElement).click(); // -> merge

    // Not re-opened: the list stays open across the trip, which is a UI
    // preference rather than a fact about the block. Toggling here would shut
    // it and the assertion below would pass on an empty list.
    const values = Array.from(pane().querySelectorAll("[data-condition]")).map((s) => (s as HTMLSelectElement).value);
    expect(values, "the control is not open").toHaveLength(3);
    expect(values, "slide 5 kept a condition set while it was the middle of another block").toEqual(["", "", ""]);

    // And what actually reaches the engine, which is where the harm would be.
    office.runMerge.mockResolvedValueOnce(OUTCOME);
    primary().click();
    await settle();
    expect((office.runMerge.mock.calls[0]?.[0] as { conditions?: unknown }).conditions).toBeUndefined();
  });

  it("keeps conditions when the numbers are retyped to the SAME block", async () => {
    /**
     * The path a careful user takes: they set a condition, go back to check the
     * slide numbers, retype the same last slide, and walk forward again. Every
     * keystroke in those boxes ran the block-moved rule, so the conditions were
     * gone — silently, with the merge button quietly offering more slides than
     * they had asked for.
     *
     * The empty box in the middle is the part that makes this hard: after the
     * first keystroke the pane names no block at all, so "has the block moved?"
     * cannot be answered from the state and every later keystroke read as a
     * move. The test types it the way a person does, one field at a time.
     */
    await toMergeWithData();
    openConditions();
    choose(5, "Last");
    backToTemplate();
    type("to", "");
    type("to", "6");
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // template -> data
    await settle();
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // data -> fields
    primary().click(); // fields -> preview
    await settle();
    (pane().querySelector("[data-forward]") as HTMLElement).click(); // -> merge

    const values = Array.from(pane().querySelectorAll("[data-condition]")).map((s) => (s as HTMLSelectElement).value);
    expect(values, "the control is not open").toHaveLength(3);
    expect(values, "the block did not move, so the condition should still be set").toEqual(["", "Last", ""]);

    // And it is what reaches the engine, which is where it matters.
    office.runMerge.mockResolvedValueOnce(OUTCOME);
    primary().click();
    await settle();
    expect((office.runMerge.mock.calls[0]?.[0] as { conditions?: unknown }).conditions).toEqual({ 5: "Last" });
  });

  it("keeps conditions when the SELECTED slides are the same block", async () => {
    /**
     * The fix for the keystroke path went in and left this one calling
     * `blockMoved` directly, so a user who set a condition, went back, and
     * pressed "use the slides I've selected" with the same slides still
     * selected lost it — the same silent widening of the merge, by the other
     * route. Two routes to one question is how the defect was written in the
     * first place; they share the rule now.
     */
    await toMergeWithData();
    openConditions();
    choose(5, "Last");
    backToTemplate();
    office.selectedBlock.mockResolvedValueOnce({ ok: true, from: 4, to: 6 });
    (pane().querySelector('[data-action="selection"]') as HTMLElement).click();
    await settle();
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // template -> data
    await settle();
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // data -> fields
    primary().click(); // fields -> preview
    await settle();
    (pane().querySelector("[data-forward]") as HTMLElement).click(); // -> merge

    const values = Array.from(pane().querySelectorAll("[data-condition]")).map((s) => (s as HTMLSelectElement).value);
    expect(values, "the control is not open").toHaveLength(3);
    expect(values, "the same slides are not a different block").toEqual(["", "Last", ""]);
  });

  it("drops them when the SELECTED slides are a different block", async () => {
    // The other direction, on the overlapping case: slide 5 is still inside the
    // new block, so a stale key would silently apply to a slide nobody set it
    // on rather than being ignored.
    await toMergeWithData();
    openConditions();
    choose(5, "Last");
    backToTemplate();
    office.selectedBlock.mockResolvedValueOnce({ ok: true, from: 3, to: 5 });
    (pane().querySelector('[data-action="selection"]') as HTMLElement).click();
    await settle();
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click();
    await settle();
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click();
    primary().click();
    await settle();
    (pane().querySelector("[data-forward]") as HTMLElement).click();

    const values = Array.from(pane().querySelectorAll("[data-condition]")).map((s) => (s as HTMLSelectElement).value);
    expect(values).toEqual(["", "", ""]);
  });

  it("keeps a condition across a new paste, and says the column is gone", async () => {
    /**
     * The opposite decision from the row filter, and deliberately: a filter is
     * about the DATA, so row 7 of a new paste is not row 7 of the old one. A
     * condition is about the TEMPLATE. Dropping it silently would rewrite the
     * user's answer to "always" and change what the merge produces with nothing
     * said — which is exactly what the engine's `unknownConditions` exists to
     * prevent.
     */
    // A condition on a column the TEMPLATE does not use, so this test is about
    // a dropped CONDITION rather than a dropped field. (It used to matter more:
    // an unmatched field blocked the merge until 2026-08-29 and this test could
    // not have reached the screen it is about. It is a caution now, so the
    // separation is for clarity rather than for reachability.)
    await reachMerge();
    backToData();
    type("paste", "First\tLast\tRegion\nAda\tLovelace\tEMEA\nGrace\tHopper\tAMER");
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // data -> fields
    primary().click(); // fields, re-reading the slides -> preview
    await settle();
    (pane().querySelector("[data-forward]") as HTMLElement).click(); // -> merge
    openConditions();
    choose(5, "Region");
    backToData();
    type("paste", "First\tLast\tCity\nAda\tLovelace\tLondon");
    office.inspectBlock.mockResolvedValueOnce(REPORT);
    primary().click(); // data -> fields
    primary().click(); // fields -> preview
    await settle();
    (pane().querySelector("[data-forward]") as HTMLElement).click(); // -> merge

    expect(said().join(" ")).toContain("No column for Region");
    const node = pane().querySelector('[data-condition="5"]') as HTMLSelectElement;
    // Still chosen, and still an option, rather than quietly reset to Always.
    expect(node.value).toBe("Region");
  });

  it("does not take a change while a merge is out", async () => {
    // Same rule as every other edit on this screen: the answer is about to be
    // written into this state and a change would make it stale.
    //
    // Asserted on what SURVIVES the run rather than on the control being
    // absent. The control now shares the merge screen with the run, so it is
    // still on screen while the call is out — and a select the user has
    // changed keeps its own value until the next draw. The draw after the
    // merge answers is where the state has its say.
    await toMergeWithData();
    openConditions();
    const held = deferred<unknown>();
    office.runMerge.mockReturnValueOnce(held.promise);
    primary().click();
    await settle();
    choose(5, "Last");
    held.resolve(OUTCOME);
    await settle();
    const node = pane().querySelector('[data-condition="5"]') as HTMLSelectElement;
    expect(node.value, "took a condition while the merge was out").toBe("");
  });
});

describe("what a blank cell does, driven", () => {
  /** To the merge step with data that has a blank in it. */
  async function toMergeWithBlanks(): Promise<void> {
    await openPane();
    await settle();
    type("from", "4");
    type("to", "6");
    office.inspectBlock.mockResolvedValueOnce({ ...REPORT, slideFields: [["First"], ["Last"]] });
    primary().click();
    await settle();
    type("paste", "First\tLast\nAda\t\nGrace\tHopper");
    primary().click(); // data -> fields
    office.inspectBlock.mockResolvedValueOnce({ ...REPORT, slideFields: [["First"], ["Last"]] });
    primary().click(); // fields -> preview
    await settle();
    (pane().querySelector("[data-forward]") as HTMLElement).click();
  }

  function choose(value: string): void {
    const node = pane().querySelector("[data-empty]") as HTMLSelectElement;
    node.value = value;
    node.dispatchEvent(new Event("input", { bubbles: true }));
  }

  it("carries the answer into the merge the user pressed", async () => {
    await toMergeWithBlanks();
    (pane().querySelector('[data-action="empties"]') as HTMLElement).click();
    choose("skip");
    await settle();

    office.runMerge.mockResolvedValueOnce({ ...OUTCOME, added: 3 });
    primary().click();
    await settle();
    expect(office.runMerge.mock.calls[0]?.[0]).toMatchObject({ onEmpty: "skip" });
  });

  it("sends nothing when the user never touched it, so an old state merges as it did", async () => {
    await toMergeWithBlanks();
    office.runMerge.mockResolvedValueOnce(OUTCOME);
    primary().click();
    await settle();
    expect(office.runMerge.mock.calls[0]?.[0]).not.toHaveProperty("onEmpty");
  });

  it("re-arms the button, because a different answer is a different merge", async () => {
    await toMergeWithBlanks();
    office.runMerge.mockResolvedValueOnce(OUTCOME);
    primary().click();
    await settle();
    expect(primary().disabled).toBe(true);

    (pane().querySelector('[data-action="empties"]') as HTMLElement).click();
    choose("keep");
    await settle();
    expect(primary().disabled, "a different merge is pressable again").toBe(false);
    expect(
      document.querySelector('.card.undo button[data-action="undo"]'),
      "and the slides that landed are still offered back",
    ).not.toBeNull();
  });

  it("keeps the keyboard on the control that was just used", async () => {
    await toMergeWithBlanks();
    (pane().querySelector('[data-action="empties"]') as HTMLElement).click();
    const node = pane().querySelector("[data-empty]") as HTMLSelectElement;
    node.focus();
    node.value = "skip";
    node.dispatchEvent(new Event("input", { bubbles: true }));
    await settle();
    expect(document.activeElement).toBe(pane().querySelector("[data-empty]"));
  });

  it("follows the slides when the user goes back and changes which fields are on them", async () => {
    /**
     * The count is read off a SNAPSHOT — the per-slide fields a template read
     * answered — and the user can go and change the slides in PowerPoint
     * between one read and the next. Nothing tells the pane that happened;
     * pressing the fields step's button is what re-reads, and the count has to
     * follow that press.
     *
     * Proven against a build where only the FIRST read fills `slideFields` in:
     * the caution then still says "1 of 2 rows will be left out" about a
     * `{{Last}}` no slide carries any more.
     */
    await toMergeWithBlanks();
    (pane().querySelector('[data-action="empties"]') as HTMLElement).click();
    choose("skip");
    await settle();
    // Ada's Last is blank and {{Last}} is on the second slide, so she goes.
    expect(pane().textContent).toContain("1 of 2 rows will be left out");

    // Back to the fields step, and this time the slides no longer carry
    // {{Last}} anywhere — the user took it off.
    (pane().querySelector("[data-back]") as HTMLElement).click(); // preview
    (pane().querySelector("[data-back]") as HTMLElement).click(); // fields
    office.inspectBlock.mockResolvedValueOnce({
      ...REPORT,
      fields: ["First"],
      slideFields: [["First"], []],
    });
    primary().click(); // re-read
    await settle();
    (pane().querySelector("[data-forward]") as HTMLElement).click(); // -> merge

    const said = pane().textContent ?? "";
    expect(said, "nothing is dropped once the blank field is off the slides").not.toContain("will be left out");
    expect(said, "and the heading goes back to one number").toContain("2 rows ×");
  });

  it("does not take a change while a merge is out", async () => {
    await toMergeWithBlanks();
    (pane().querySelector('[data-action="empties"]') as HTMLElement).click();
    const held = deferred<unknown>();
    office.runMerge.mockReturnValueOnce(held.promise);
    primary().click();
    await settle();
    choose("skip");
    held.resolve(OUTCOME);
    await settle();
    expect((pane().querySelector("[data-empty]") as HTMLSelectElement).value).toBe("blank");
  });
});

describe("the run record while the host has not answered", () => {
  /**
   * The case the record exists for, driven through the real `main.ts`.
   *
   * `render.ts` gated the record on the run being over and `main.ts` wrote it
   * once, in a `finally`. Between them, a host that never answers produced a
   * pane showing "Waiting on PowerPoint: …" and nothing else — no record, on
   * the one run somebody needs to explain. A task pane has no devtools and
   * cannot hand over a file, so what is on screen is the whole channel.
   *
   * Asserted here rather than only in `pane-render.test.ts` because that file
   * hands `render` a state with a `log` already in it. The question this one
   * asks is whether anything PUTS one there while the call is out.
   */
  it("is on screen, with the call named, before the merge returns", async () => {
    await reachMerge();
    const held = deferred<unknown>();
    office.runMerge.mockReturnValueOnce(held.promise);
    primary().click();
    await settle();

    const log = pane().querySelector(".runlog pre")?.textContent ?? "";
    expect(log, "no record while the host has not answered").toContain("=== SSF MERGE RUN LOG ===");
    // The environment line the run emits after its mark, so the record is
    // useful and not merely present.
    expect(log).toContain("run starting");
    expect(pane().querySelector(".runlog summary")?.textContent).toBe("What this run has done so far");
    // The "Waiting on PowerPoint: …" line is deliberately NOT asserted here.
    // It comes from `inFlight`, which is set by a host call tracing "issued" —
    // and this harness mocks `runMerge`, so no host call happens and the line
    // cannot appear. Asserting it would fail against correct code, and
    // asserting its absence would pin a fact about the mock rather than about
    // the pane. `pane-render.test.ts` covers that line directly.

    held.resolve(OUTCOME);
    await settle();
    // And it is still there afterwards, now claiming to be finished.
    expect(pane().querySelector(".runlog summary")?.textContent).toBe("What this run did, step by step");
  });

  it("starts empty for each run rather than showing the last one", async () => {
    // `beginRun` clears the entries and the state's log goes with it. Without
    // that, the first thing on screen during a merge is the PREVIOUS merge's
    // record, which reads as this one having already got that far.
    await reachMerge();
    office.runMerge.mockResolvedValueOnce({ ...OUTCOME, detail: "first" });
    primary().click();
    await settle();
    const first = pane().querySelector(".runlog pre")?.textContent ?? "";
    expect(first).toContain("run starting");

    // A second run, held open, must not be showing the first one's lines.
    (pane().querySelector('[data-back="preview"]') as HTMLElement).click();
    (pane().querySelector("[data-forward]") as HTMLElement).click();
    const held = deferred<unknown>();
    office.runMerge.mockReturnValueOnce(held.promise);
    primary().click();
    await settle();
    const second = pane().querySelector(".runlog pre")?.textContent ?? "";
    expect(second.split("\n").length, "the second run shows the first one's lines").toBeLessThan(
      first.split("\n").length + 3,
    );
    held.resolve(OUTCOME);
    await settle();
  });
});

describe("a template with no fields on it yet", () => {
  /**
   * The order this pane was reordered for, driven through the real `main.ts`.
   *
   * The state it starts from is the one that was reported: a fresh deck, two
   * empty slides, and a user who cannot name a single placeholder because the
   * names are their data's column headers and the data is not attached yet.
   * The old order refused at step 1 and told them to go and type names nobody
   * had; the fix is that the slides come first, the data second, and the
   * fields are put on the slides from a button.
   */
  async function toFields(): Promise<void> {
    await openPane();
    await settle();
    type("from", "2");
    type("to", "3");
    // The template read ANSWERS for a block with nothing on it now. `runMerge`
    // still refuses one — see `allowEmpty` — and the pane refuses it too,
    // before a host call is spent.
    office.inspectBlock.mockResolvedValueOnce({ ok: true, detail: "0 placeholders in slides 2 to 3.", fields: [] });
    primary().click();
    await settle();
    type("paste", "First\tCity\nAda\tLondon");
    primary().click(); // data -> fields
  }

  it("goes on to the data step rather than refusing", async () => {
    await openPane();
    await settle();
    type("from", "2");
    type("to", "3");
    office.inspectBlock.mockResolvedValueOnce({ ok: true, detail: "0 placeholders in slides 2 to 3.", fields: [] });
    primary().click();
    await settle();
    // The step label is uppercased by CSS, so the text node is title case.
    expect(document.body.textContent).toContain("Step 2 of 5");
    expect(field("paste"), "nothing to paste into").not.toBeNull();
  });

  it("offers a button per column, spelled the way the engine reads it", async () => {
    await toFields();
    const chips = Array.from(pane().querySelectorAll("[data-insert]"));
    expect(chips.map((c) => c.getAttribute("data-insert"))).toEqual(["First", "City"]);
    // The TOKEN on the face, because that is what lands on the slide.
    expect(chips.map((c) => c.textContent)).toEqual(["{{First}}", "{{City}}"]);
  });

  it("puts the field where the cursor is", async () => {
    await toFields();
    office.insertTextAtCursor.mockResolvedValueOnce({ ok: true });
    (pane().querySelector('[data-insert="City"]') as HTMLElement).click();
    await settle();
    expect(office.insertTextAtCursor).toHaveBeenCalledWith("{{City}}");
    // Says what to do NEXT: the insert lands on the slide, not in the pane, so
    // without this the fields list stays empty and nothing says why.
    expect(document.body.textContent).toContain("Check the slides");
  });

  it("falls back to the clipboard when the host will not type it in", async () => {
    // The reason the control is worth building at all. A refusal here is
    // ordinary — `setSelectedDataAsync` needs an insertion point, and a user
    // who has not clicked into a text box has none.
    const copied: string[] = [];
    Object.defineProperty(globalThis.navigator, "clipboard", {
      configurable: true,
      value: {
        writeText: (t: string) => {
          copied.push(t);
          return Promise.resolve();
        },
      },
    });
    await toFields();
    office.insertTextAtCursor.mockResolvedValueOnce({ ok: false, why: "no insertion point" });
    (pane().querySelector('[data-insert="First"]') as HTMLElement).click();
    await settle();
    expect(copied, "nothing reached the clipboard").toEqual(["{{First}}"]);
    expect(document.body.textContent).toContain("clipboard");
    // The host's own words, because they are what separates "click into a text
    // box first" from "this host cannot do it at all".
    expect(document.body.textContent).toContain("no insertion point");
  });

  it("names the token when even the clipboard is refused", async () => {
    // A task pane is a nested cross-origin iframe, where `navigator.clipboard`
    // is gated on a permission the host may not have granted. All three
    // outcomes name the token; none of them is silence.
    Object.defineProperty(globalThis.navigator, "clipboard", {
      configurable: true,
      value: { writeText: () => Promise.reject(new Error("denied")) },
    });
    await toFields();
    office.insertTextAtCursor.mockResolvedValueOnce({ ok: false, why: "no insertion point" });
    (pane().querySelector('[data-insert="First"]') as HTMLElement).click();
    await settle();
    expect(document.body.textContent).toContain("{{First}}");
    expect(document.body.textContent).toContain("by hand");
  });

  it("reads the slides again when the primary is pressed, and goes on", async () => {
    // Nothing tells this pane that the user typed on a slide — there is no
    // document-changed event for slide text — so the fields it lists are as old
    // as the last read. The step's own primary is that read.
    await toFields();
    expect(office.inspectBlock).toHaveBeenCalledTimes(1);
    office.inspectBlock.mockResolvedValueOnce({ ok: true, detail: "2 placeholders.", fields: ["First", "City"] });
    primary().click();
    await settle();
    expect(office.inspectBlock).toHaveBeenCalledTimes(2);
    expect(document.body.textContent).toContain("Step 4 of 5");
  });

  it("will not preview or merge a block that still has no fields", async () => {
    // The engine refuses this too, and must — N identical copies is never what
    // anybody meant. Said here so the refusal costs no host call.
    await toFields();
    office.inspectBlock.mockResolvedValueOnce({ ok: true, detail: "0 placeholders.", fields: [] });
    primary().click();
    await settle();
    // Still on the fields step, and said why.
    expect(document.body.textContent).toContain("Step 3 of 5");
    expect(pane().textContent).toContain("carry no fields yet");
    expect(office.runMerge).not.toHaveBeenCalled();
  });
});

describe("a run the pane never came back from", () => {
  /**
   * The crumb's whole reason, driven through the real `main.ts`.
   *
   * `merge()` writes `{added: 0}` BEFORE handing the package to PowerPoint,
   * because a tab that dies during that call never comes back to write the real
   * number. `readCrumb` refused zero, so the record was write-only in exactly
   * that window — readable only after the run it was insurance against had
   * already succeeded, which is the one case it is not needed for.
   *
   * What it may do with it is bounded. Nothing here knows how many slides
   * landed: taking the deck's growth as the answer would sweep whatever has
   * been appended since, which `sweepPlan` refuses by design. So the user is
   * TOLD, with both numbers and where to look, and never offered a delete.
   */
  const KEY = "ssf-merge.run.v1";
  // The deck these crumbs belong to. The store is shared by every deck on the
  // add-in's origin, so a crumb names the one it was written on and the pane
  // only answers for a match — this is what the mocked `documentKey` returns.
  const DECK = "https://example-my.sharepoint.com/personal/x/Documents/deck.pptx";

  // The store survives `vi.resetModules`, and earlier tests in this file run
  // real merges that drop crumbs of their own.
  beforeEach(() => localStorage.removeItem(KEY));

  it("says a merge did not finish, with the deck's size either side of it", async () => {
    localStorage.setItem(
      KEY,
      JSON.stringify({
        kind: "ssf-merge-run",
        deckAtStart: 4,
        added: 0,
        runId: "pending",
        startedAt: "2026-08-27T10:00:00.000Z",
        doc: DECK,
      }),
    );
    office.slideCount.mockReset().mockResolvedValue(16);
    await openPane();
    await settle();

    const text = pane().textContent ?? "";
    expect(text, "said nothing about the run that died").toContain("did not finish");
    expect(text).toContain("4 slide(s) before it");
    expect(text).toContain("16");
    // Told, never offered. There is deliberately no assertion here that the
    // undo card is absent: it is drawn on the merge step and this is the
    // template step, so it would pass whatever the branch did — the vacuous
    // shape this repo keeps finding. What makes it true is that the branch
    // sets no `added` and no `last`, which is `crumb.test.ts`'s to check and
    // `sweepPlan`'s to enforce.
  });

  it("says it once, because there is no action attached to it", async () => {
    localStorage.setItem(
      KEY,
      JSON.stringify({
        kind: "ssf-merge-run",
        deckAtStart: 4,
        added: 0,
        runId: "pending",
        startedAt: "2026-08-27",
        doc: DECK,
      }),
    );
    office.slideCount.mockReset().mockResolvedValue(16);
    await openPane();
    await settle();
    expect(pane().textContent).toContain("did not finish");

    // A second open. Left in the store the same sentence would greet the user
    // on every open for ever, about a run they have already dealt with.
    await openPane();
    await settle();
    expect(pane().textContent, "repeats on every open").not.toContain("did not finish");
  });

  it("still OFFERS the way back when the run got as far as counting", async () => {
    // The other branch: a crumb carrying a real count can be swept.
    //
    // "Offers" was a claim about a sentence until this test pressed the
    // button. The branch set `added`, `deckAtStart` and `last`, so everything
    // an undo needs was in hand — and the card is drawn on the MERGE step,
    // which a pane that has just opened cannot reach: it takes a template
    // read, a paste and a field check to get there, none of which has anything
    // to do with taking last night's slides back out.
    localStorage.setItem(
      KEY,
      JSON.stringify({
        kind: "ssf-merge-run",
        deckAtStart: 12,
        added: 6,
        runId: "r1",
        startedAt: "2026-08-27",
        doc: DECK,
      }),
    );
    office.slideCount.mockReset().mockResolvedValue(18);
    await openPane();
    await settle();
    expect(pane().textContent).toContain("added 6 slides");

    const undo = document.querySelector<HTMLButtonElement>('.card.undo button[data-action="undo"]');
    expect(undo, "a way back, not only a sentence about one").not.toBeNull();
    expect(pane().textContent).toContain("Remove slides 13 to 18");

    office.undoMerge.mockResolvedValueOnce({ removed: 6, detail: "removed 6 slide(s) from index 12" });
    // The deck's size is ASKED for after a sweep rather than computed from the
    // pane's cached number, which is stale exactly when the user has been
    // editing by hand.
    office.slideCount.mockResolvedValueOnce(12);
    undo?.click();
    await settle();
    expect(office.undoMerge.mock.calls[0]?.[0]).toMatchObject({ deckAtStart: 12, added: 6 });
    expect(pane().textContent).toContain("Your deck holds 12");
  });

  it("does not move the user who has already started typing", async () => {
    // The deck count resolves a second or two after the pane opens, and the
    // crumb is read when it does. Jumping to the last step then would take
    // somebody mid-keystroke off the box they are in — the same hazard the
    // caret restore exists for, one level up.
    localStorage.setItem(
      KEY,
      JSON.stringify({
        kind: "ssf-merge-run",
        deckAtStart: 12,
        added: 6,
        runId: "r1",
        startedAt: "2026-08-27",
        doc: DECK,
      }),
    );
    const count = deferred<number>();
    office.slideCount.mockReset().mockReturnValue(count.promise);
    await openPane();
    type("from", "4");
    count.resolve(18);
    await settle();

    expect(pane().textContent, "still on the step the user was typing in").toContain("Which slides repeat?");
    expect(pane().textContent, "and still told about the run").toContain("added 6 slides");
  });
});

describe("the pictures the user picked", () => {
  const PHOTO_REPORT = { ...REPORT, fields: ["First", "Photo"] };

  /** A picker choice. Shaped as the pane reads it: `length`, iterable, `name`, `arrayBuffer`. */
  function choose(...files: { name: string; bytes?: Uint8Array; refuse?: boolean }[]): void {
    const node = field("images") as HTMLInputElement;
    const list = files.map((f) => ({
      name: f.name,
      arrayBuffer: () =>
        f.refuse
          ? Promise.reject(new Error("the file has moved"))
          : Promise.resolve((f.bytes ?? new Uint8Array([1, 2, 3])).buffer),
    }));
    Object.defineProperty(node, "files", { configurable: true, value: list });
    node.dispatchEvent(new Event("input", { bubbles: true }));
  }

  /** To the data step, with data whose cells name pictures. */
  async function reachData(): Promise<void> {
    await openPane();
    await settle();
    type("from", "4");
    type("to", "6");
    office.inspectBlock.mockResolvedValueOnce(PHOTO_REPORT);
    primary().click();
    await settle();
    type("paste", "First\tPhoto\nAda\tada.png\nGrace\tgrace.png");
  }

  it("keeps the first folder's pictures when a second folder is picked", async () => {
    /**
     * A browser's picker returns one directory's selection, and a spreadsheet
     * built from a photo library routinely names files in several. Picking the
     * second folder replaced the first: the tally then reported every name
     * from folder one as missing, with the files sitting on the disk the
     * author had just chosen them from and no way at all to attach both.
     */
    await reachData();
    choose({ name: "ada.png", bytes: new Uint8Array([1]) });
    await settle();
    choose({ name: "grace.png", bytes: new Uint8Array([2]) });
    await settle();

    expect(pane().textContent, "both names are answered for").toContain("All 2 pictures matched.");
    // And the reader is told once that the picker added rather than replaced.
    expect(pane().textContent).toContain("Added to the pictures already attached — 2 files now.");
  });

  it("hands them to the merge, keyed by the name the file has on disk", async () => {
    await reachData();
    choose({ name: "ada.png", bytes: new Uint8Array([1]) }, { name: "grace.png", bytes: new Uint8Array([2]) });
    await settle();
    primary().click(); // data -> fields
    office.inspectBlock.mockResolvedValueOnce(PHOTO_REPORT);
    primary().click(); // fields -> preview
    await settle();
    (pane().querySelector("[data-forward]") as HTMLElement).click();
    office.runMerge.mockResolvedValueOnce(OUTCOME);
    primary().click();
    await settle();

    const request = office.runMerge.mock.calls[0]?.[0] as { images?: Map<string, Uint8Array> };
    expect([...(request.images?.keys() ?? [])]).toEqual(["ada.png", "grace.png"]);
    expect(request.images?.get("ada.png")).toEqual(new Uint8Array([1]));
  });

  it("counts a file it could not read and keeps the rest", async () => {
    // A file moved, renamed or on a disconnected drive between the dialog and
    // the read. Throwing there would lose the pictures that ARE in hand.
    await reachData();
    choose({ name: "ada.png" }, { name: "gone.png", refuse: true });
    await settle();

    expect(said().join(" ")).toContain("1 of the 2 file(s) could not be read");
    expect(pane().textContent).toContain("1 of 2 matched. Missing: grace.png.");
  });

  it("takes the choice even while a host call is out", async () => {
    // Reading files touches no host, so it cannot go stale against an answer in
    // flight — and refusing it would leave the browser's own dialog having
    // visibly taken a choice the pane ignored, with nothing said.
    //
    // The picker is SYNTHESISED here rather than clicked, because the data step
    // draws no host call of its own and so never shows this control while one
    // is out. The subject is the ordering inside the delegated handler, which
    // is what a later step's control on this screen would depend on.
    await reachData();
    primary().click(); // data -> fields, no host call of its own
    await settle();
    const held = deferred<unknown>();
    office.inspectBlock.mockReturnValueOnce(held.promise);
    primary().click(); // the fields step re-reads the slides: held open
    await settle();
    expect(primary().disabled, "the pane is running").toBe(true);

    const node = document.createElement("input");
    node.type = "file";
    node.setAttribute("data-field", "images");
    pane().append(node);
    Object.defineProperty(node, "files", {
      configurable: true,
      value: [{ name: "ada.png", arrayBuffer: () => Promise.resolve(new Uint8Array([1]).buffer) }],
    });
    node.dispatchEvent(new Event("input", { bubbles: true }));
    await settle();

    held.resolve(PHOTO_REPORT);
    await settle();
    (pane().querySelector("[data-back]") as HTMLElement).click(); // preview -> fields
    (pane().querySelector("[data-back]") as HTMLElement).click(); // fields -> data
    expect(pane().textContent).toContain("1 of 2 matched. Missing: grace.png.");
  });
});

describe("a placeholder with no column", () => {
  it("warns on the merge screen and still merges", async () => {
    /**
     * The whole point of turning that gate into a caution, end to end.
     *
     * Until 2026-08-29 this state stopped the merge, which is also what made
     * step 4 a dead end — the only way out of the preview step was drawn on the
     * merge being reachable. The engine has always left such a placeholder on
     * the slide, the preview step has always run the ordinary merge with one,
     * and docs/MANUAL.md promises it.
     */
    await openPane();
    await settle();
    type("from", "4");
    type("to", "6");
    office.inspectBlock.mockResolvedValueOnce({
      ...REPORT,
      fields: ["First", "Nickname"],
      detail: "2 placeholders in slides 4 to 6.",
    });
    primary().click(); // template -> data
    await settle();
    // No Nickname column, deliberately. This is the test kit's own shape.
    type("paste", "First\tLast\nAda\tLovelace\nGrace\tHopper");
    primary().click(); // data -> fields
    office.inspectBlock.mockResolvedValueOnce({ ...REPORT, fields: ["First", "Nickname"] });
    primary().click(); // fields -> preview
    await settle();

    (pane().querySelector("[data-forward]") as HTMLElement).click(); // -> merge

    expect(said().join(" "), "the user is told which one").toContain("No column for Nickname");
    expect(said().join(" ")).toContain("stay on the slides as written");
    expect(primary().disabled, "the caution became a wall again").toBe(false);

    office.runMerge.mockResolvedValueOnce(OUTCOME);
    primary().click();
    await settle();
    expect(office.runMerge).toHaveBeenCalledTimes(1);
    expect(said().join(" ")).toContain("6 slides added");
  });
});
