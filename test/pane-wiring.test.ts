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
function primary(): HTMLButtonElement {
  return pane().querySelector("button.primary") as HTMLButtonElement;
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
    office.slideCount.mockResolvedValueOnce(18); // 12 before, 18 after: six landed
    office.runMerge.mockRejectedValueOnce(new Error("gave up waiting for: inserting the merged deck"));
    primary().click();
    await settle();

    const told = said().join(" ");
    expect(told).toContain("6 slide");
    expect(told).toContain("landed anyway");
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
    expect(primary().textContent).toBe("Remove the preview");
  });

  it("takes it back with the same clamped sweep an undo uses", async () => {
    await reachPreview();
    office.runMerge.mockResolvedValueOnce(PREVIEW);
    primary().click();
    await settle();

    office.undoMerge.mockResolvedValueOnce({ removed: 3, detail: "removed 3 slide(s) from index 12" });
    primary().click();
    await settle();

    expect(office.undoMerge).toHaveBeenCalledWith(expect.objectContaining({ deckAtStart: 12, added: 3 }));
    expect(pane().querySelectorAll(".card.undo")).toHaveLength(0);
    expect(primary().textContent).toBe("Preview the first row");
  });

  it("says so when the sweep left some of it behind", async () => {
    // A sweep that removed fewer slides than it asked for leaves part of the
    // preview in the deck, and the user is the only one who can finish it.
    await reachPreview();
    office.runMerge.mockResolvedValueOnce(PREVIEW);
    primary().click();
    await settle();

    office.undoMerge.mockResolvedValueOnce({ removed: 1, detail: "asked for 3 slide(s) and the deck shrank by 1" });
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
    primary().click();
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
  async function afterMerge(): Promise<HTMLElement> {
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
    office.runMerge.mockResolvedValueOnce(OUTCOME);
    primary().click();
    await settle();
    return root;
  }

  const undoButton = (): HTMLButtonElement | null =>
    document.querySelector<HTMLButtonElement>('.card.undo button[data-action="undo"]');

  it("offers the way back once a merge has landed", async () => {
    // `undoInsert` and `sweepPlan` were built and tested before this and were
    // reachable from nothing — the numbers were kept, the sentence was
    // written, and no view rendered either.
    await afterMerge();
    expect(undoButton()).not.toBeNull();
  });

  it("sweeps with the run's OWN numbers, not the pane's current ones", async () => {
    // The clamps are only worth anything against the count taken before the
    // run inserted. Handing the sweep a number the pane happens to hold now is
    // how a positional delete reaches slides the user owned first.
    await afterMerge();
    office.undoMerge.mockResolvedValueOnce({ removed: 6, detail: "removed 6 slide(s) from index 12" });
    undoButton()?.click();
    await settle();

    expect(office.undoMerge).toHaveBeenCalledTimes(1);
    expect(office.undoMerge.mock.calls[0]?.[0]).toMatchObject({ deckAtStart: 12, added: 6 });
  });

  it("puts the way back away once the slides are gone", async () => {
    await afterMerge();
    office.undoMerge.mockResolvedValueOnce({ removed: 6, detail: "removed 6 slide(s) from index 12" });
    undoButton()?.click();
    await settle();
    expect(undoButton()).toBeNull();
    expect(document.body.textContent).toContain("back to 12");
  });

  it("KEEPS the way back when the sweep only got some of them", async () => {
    // A partial sweep leaves slides in the deck and the user is the only one
    // who can finish the job, so the button has to stay.
    await afterMerge();
    office.undoMerge.mockResolvedValueOnce({ removed: 2, detail: "asked for 6 and the deck shrank by 2" });
    undoButton()?.click();
    await settle();
    expect(undoButton(), "still offered").not.toBeNull();
    expect(document.body.textContent).toContain("Some of the merge is still there");
  });

  it("says so when the sweep refused, and leaves the deck alone", async () => {
    // `sweepPlan` refuses when the deck gained more than the run added,
    // because the last N slides are then somebody else's. That refusal must
    // reach the user as a sentence rather than as a silent no-op.
    await afterMerge();
    office.undoMerge.mockResolvedValueOnce({ removed: 0, detail: "nothing to take back (deck was 12, is 20)" });
    undoButton()?.click();
    await settle();
    expect(document.body.textContent).toContain("Nothing was removed");
    expect(undoButton(), "still offered, because nothing went").not.toBeNull();
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
    office.slideCount.mockResolvedValueOnce(18); // 12 before, 18 after
    office.runMerge.mockRejectedValueOnce(new Error("gave up waiting for: inserting the merged deck"));
    primary().click();
    await settle();

    expect(document.body.textContent).toContain("landed anyway");
    expect(undoButton(), "six slides in the deck and a way to remove them").not.toBeNull();
    office.undoMerge.mockResolvedValueOnce({ removed: 6, detail: "removed 6 slide(s) from index 12" });
    undoButton()?.click();
    await settle();
    expect(office.undoMerge.mock.calls[0]?.[0]).toMatchObject({ deckAtStart: 12, added: 6 });
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
      JSON.stringify({ kind: "ssf-merge-run", deckAtStart: 4, added: 0, runId: "pending", startedAt: "2026-08-27" }),
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
      JSON.stringify({ kind: "ssf-merge-run", deckAtStart: 12, added: 6, runId: "r1", startedAt: "2026-08-27" }),
    );
    office.slideCount.mockReset().mockResolvedValue(18);
    await openPane();
    await settle();
    expect(pane().textContent).toContain("added 6 slide(s)");

    const undo = document.querySelector<HTMLButtonElement>('.card.undo button[data-action="undo"]');
    expect(undo, "a way back, not only a sentence about one").not.toBeNull();
    expect(pane().textContent).toContain("Remove slides 13 to 18");

    office.undoMerge.mockResolvedValueOnce({ removed: 6, detail: "removed 6 slide(s) from index 12" });
    undo?.click();
    await settle();
    expect(office.undoMerge.mock.calls[0]?.[0]).toMatchObject({ deckAtStart: 12, added: 6 });
    expect(pane().textContent).toContain("back to 12");
  });

  it("does not move the user who has already started typing", async () => {
    // The deck count resolves a second or two after the pane opens, and the
    // crumb is read when it does. Jumping to the last step then would take
    // somebody mid-keystroke off the box they are in — the same hazard the
    // caret restore exists for, one level up.
    localStorage.setItem(
      KEY,
      JSON.stringify({ kind: "ssf-merge-run", deckAtStart: 12, added: 6, runId: "r1", startedAt: "2026-08-27" }),
    );
    const count = deferred<number>();
    office.slideCount.mockReset().mockReturnValue(count.promise);
    await openPane();
    type("from", "4");
    count.resolve(18);
    await settle();

    expect(pane().textContent, "still on the step the user was typing in").toContain("Which slides repeat?");
    expect(pane().textContent, "and still told about the run").toContain("added 6 slide(s)");
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
