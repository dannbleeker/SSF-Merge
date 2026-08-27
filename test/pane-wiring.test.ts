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
  inspectBlock: vi.fn<(r: { from: number; to: number }) => Promise<unknown>>(),
  runMerge: vi.fn<(r: unknown) => Promise<unknown>>(),
  undoMerge: vi.fn<(o: unknown) => Promise<unknown>>(),
}));

vi.mock("../src/office/powerpoint.js", () => ({
  ready: office.ready,
  slideCount: office.slideCount,
  selectedBlock: office.selectedBlock,
  canReadSelection: office.canReadSelection,
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

let onReady: () => void;

/** Load the pane fresh, run its Office.onReady, and hand back the root. */
async function openPane(): Promise<HTMLElement> {
  document.body.innerHTML = '<div id="pane"></div>';
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

/** Walk from a fresh pane to the merge step with a block and data in hand. */
async function reachMerge(): Promise<void> {
  await openPane();
  await settle();
  type("from", "4");
  type("to", "6");
  office.inspectBlock.mockResolvedValueOnce(REPORT);
  primary().click();
  await settle();
  type("paste", "First\tLast\nAda\tLovelace\nGrace\tHopper");
  primary().click(); // fields -> preview
  // The preview step's primary SHOWS a row rather than advancing, so the way
  // forward is the link beside it.
  (pane().querySelector("[data-forward]") as HTMLElement).click();
}

beforeEach(() => {
  office.slideCount.mockReset().mockResolvedValue(12);
  office.inspectBlock.mockReset();
  office.runMerge.mockReset();
  office.undoMerge.mockReset();
  office.selectedBlock.mockReset();
  office.canReadSelection.mockReset().mockReturnValue(true);
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
    // twice to the paste box, change the data, and the button is live again.
    (pane().querySelector("[data-back]") as HTMLElement).click(); // preview
    (pane().querySelector("[data-back]") as HTMLElement).click(); // fields
    type("paste", "First\tLast\nAda\tLovelace");
    primary().click(); // fields -> preview
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
    primary().click();
    await settle();
    type("paste", "First\tLast\nAda\tLovelace\nGrace\tHopper");
    primary().click(); // fields -> preview
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
    // is an EXTRA. It shipped unguarded — the call went in on 174 rounds of
    // evidence that it is not wedged, without anyone asking which version
    // introduced it. Safe to call and present are different questions.
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
    primary().click(); // fields -> preview
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
    type("paste", "First\tLast\nZoe\tZed\nYan\tYates");
    primary().click();
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
