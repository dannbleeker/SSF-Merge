// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { render } from "../src/pane/render.js";
import type { PaneState } from "../src/pane/steps.js";

const ready: PaneState = {
  block: { from: 4, to: 6 },
  fields: ["First", "Last"],
  columns: ["First", "Last"],
  rows: 240,
  previewing: false,
  deckSize: 12,
};

function paneFor(state: PaneState, step: Parameters<typeof render>[2]) {
  const root = document.createElement("div");
  render(root, state, step);
  return root;
}

describe("the pane's chrome", () => {
  it("says which step of how many, because that is what a first-timer asks", () => {
    expect(paneFor(ready, "fields").querySelector(".step-of")?.textContent).toBe("Step 2 of 4 · Fields");
  });

  it("draws a four-segment rail whatever the step", () => {
    expect(paneFor(ready, "template").querySelectorAll(".rail li")).toHaveLength(4);
  });

  it("puts exactly one primary button on the screen, and it is last", () => {
    const main = paneFor(ready, "merge").querySelector("main");
    const primaries = main?.querySelectorAll("button.primary") ?? [];
    expect(primaries).toHaveLength(1);
    expect(main?.lastElementChild).toBe(primaries[0]);
  });
});

describe("the orange budget", () => {
  it("spends it on the tick in the ordinary case", () => {
    const pane = paneFor(ready, "merge");
    expect(pane.querySelectorAll(".tick")).toHaveLength(1);
    expect(pane.querySelectorAll(".card.undo")).toHaveLength(0);
  });

  it("never puts two oranges on one screen, over every state and every step", () => {
    // The rule the layout was approved on, and the one jsdom cannot see: it has
    // no colour, so "orange" here is counted by the classes that carry it. A
    // screenshot found the fields step drawing the tick AND an orange-bordered
    // chip; this is that finding turned into something CI can hold.
    const states: PaneState[] = [
      { fields: [], previewing: false },
      ready,
      { ...ready, previewing: true },
      { ...ready, fields: ["First", "Nickname"] },
      { ...ready, fields: ["First", "Nickname"], previewing: true },
      { ...ready, rows: 1, block: { from: 4, to: 4 } },
    ];
    for (const state of states) {
      for (const step of ["template", "fields", "preview", "merge"] as const) {
        const pane = paneFor(state, step);
        const oranges =
          pane.querySelectorAll(".tick").length +
          pane.querySelectorAll(".card.undo").length +
          pane.querySelectorAll('.fields li[data-matched="no"]').length;
        expect(oranges, `${step} with ${JSON.stringify(state)}`).toBeLessThanOrEqual(1);
      }
    }
  });

  it("moves it to the warning while a preview is showing, and drops the tick", () => {
    // Two oranges in one glance and neither means anything. This is the rule
    // the layout was approved on, and it is invisible to a screenshot of one
    // state — only the pair of renders shows it.
    const pane = paneFor({ ...ready, previewing: true }, "merge");
    expect(pane.querySelectorAll(".tick")).toHaveLength(0);
    expect(pane.querySelectorAll(".card.undo")).toHaveLength(1);
  });
});

describe("what the buttons say", () => {
  it("carries the number into the label", () => {
    expect(paneFor(ready, "merge").querySelector("button.primary")?.textContent).toBe("Add 720 slides");
  });

  it("disables the primary and says why when the step is blocked", () => {
    const pane = paneFor({ fields: [], previewing: false }, "merge");
    expect(pane.querySelector("button.primary")?.hasAttribute("disabled")).toBe(true);
    expect(pane.querySelector(".blocked")?.textContent).toContain("repeat");
  });
});

describe("the merge screen", () => {
  it("does not state the same arithmetic twice", () => {
    // The heading carries "240 rows x 3 slides" and the card used to repeat it
    // verbatim, which reads as a rendering bug rather than emphasis.
    const pane = paneFor(ready, "merge");
    const text = pane.textContent ?? "";
    expect(text.split("240 rows × 3 slides")).toHaveLength(2);
  });
});

describe("the field list", () => {
  it("marks a placeholder with no column, and names it in the card", () => {
    const pane = paneFor({ ...ready, fields: ["First", "Nickname"] }, "fields");
    const marked = Array.from(pane.querySelectorAll(".fields li")).filter(
      (li) => li.getAttribute("data-matched") === "no",
    );
    expect(marked.map((li) => li.textContent)).toEqual(["Nickname"]);
    expect(pane.querySelector(".card")?.textContent).toContain("Nickname");
  });
});

describe("text that came from a file somebody pasted", () => {
  it("is written as TEXT, never as markup", () => {
    // A column header is user data. innerHTML here would put a script tag from
    // a spreadsheet into the pane, which runs with the add-in's own privileges.
    const nasty = "<img src=x onerror=alert(1)>";
    const pane = paneFor({ ...ready, fields: [nasty], columns: [] }, "fields");
    expect(pane.querySelector(".fields li")?.textContent).toBe(nasty);
    expect(pane.querySelectorAll("img")).toHaveLength(0);
  });
});

describe("re-rendering", () => {
  it("replaces the pane rather than adding to it", () => {
    // A stale count left behind by a half-clear is a button that says 720 and
    // adds 30.
    const root = document.createElement("div");
    render(root, ready, "merge");
    render(root, { ...ready, rows: 10 }, "merge");
    expect(root.querySelectorAll("button.primary")).toHaveLength(1);
    expect(root.querySelector("button.primary")?.textContent).toBe("Add 30 slides");
  });
});
