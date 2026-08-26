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
      // The states the controls added. A budget checked over the states that
      // existed when it was written is a budget that stops covering the pane
      // the first time the pane grows.
      { ...ready, draft: { from: "6", to: "4" } },
      { ...ready, draft: { from: "6", to: "4" }, previewing: true },
      { ...ready, paste: "First\tLast", rows: 0, columns: undefined },
      { ...ready, notice: "PowerPoint would not name every slide." },
      { ...ready, notice: "PowerPoint would not name every slide.", previewing: true },
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

describe("counting placeholders", () => {
  it("says 1 placeholder, not 1 placeholders", () => {
    // The zero case was special-cased and the ONE case was not, so a name badge
    // or certificate template holding a single {{Name}} announced
    // "1 placeholders". The screenshot script only ever renders three, so
    // nothing showed it either.
    expect(paneFor({ ...ready, fields: ["Name"] }, "fields").querySelector("h1")?.textContent).toBe("1 placeholder");
  });

  it("still says the plural for two and the sentence for none", () => {
    expect(paneFor({ ...ready, fields: ["A", "B"] }, "fields").querySelector("h1")?.textContent).toBe("2 placeholders");
    expect(paneFor({ ...ready, fields: [] }, "fields").querySelector("h1")?.textContent).toBe("No placeholders found");
  });
});

describe("the two slide-number boxes", () => {
  it("puts both on the template step, named for main.ts to read", () => {
    const pane = paneFor({ fields: [], previewing: false }, "template");
    expect(pane.querySelector('input[data-field="from"]')).not.toBeNull();
    expect(pane.querySelector('input[data-field="to"]')).not.toBeNull();
  });

  it("sets the box's VALUE, not its value attribute", () => {
    // setAttribute("value", …) sets the default an input reverts to. The pane
    // re-renders on every keystroke, so with the attribute the box would snap
    // back to what it held before the key that caused the render.
    const pane = paneFor({ fields: [], previewing: false, draft: { from: "4", to: "6" } }, "template");
    const from = pane.querySelector('input[data-field="from"]');
    expect(from).toBeInstanceOf(HTMLInputElement);
    expect((from as HTMLInputElement).value).toBe("4");
  });

  it("shows what is wrong with the boxes, and only once there is something", () => {
    const half = paneFor({ fields: [], previewing: false, draft: { from: "4", to: "" } }, "template");
    expect(half.querySelectorAll(".blocked")).toHaveLength(0);
    const wrong = paneFor({ fields: [], previewing: false, draft: { from: "6", to: "4" } }, "template");
    expect(wrong.querySelector(".blocked")?.textContent).toContain("ends before it starts");
  });

  it("carries the typed numbers into the heading and the button", () => {
    const pane = paneFor({ fields: [], previewing: false, draft: { from: "2", to: "5" } }, "template");
    expect(pane.querySelector("h1")?.textContent).toContain("Slides 2 to 5");
    expect(pane.querySelector("button.primary")?.textContent).toBe("Use slides 2 to 5");
  });
});

describe("the paste box", () => {
  it("is on the fields step, named for main.ts to read", () => {
    expect(paneFor(ready, "fields").querySelector('textarea[data-field="paste"]')).not.toBeNull();
  });

  it("holds what was pasted as its VALUE, so a re-render does not empty it", () => {
    const pane = paneFor({ ...ready, paste: "First\tLast\nAda\tLovelace" }, "fields");
    const box = pane.querySelector("textarea");
    expect((box as HTMLTextAreaElement).value).toContain("Ada");
  });

  it("names the COLUMNS it found, not just a row count", () => {
    // A paste that came through as plain text parses into one column, and a
    // row count alone looks perfectly healthy when that happens.
    const text = paneFor({ ...ready, columns: ["First", "Last"], rows: 240 }, "fields").textContent ?? "";
    expect(text).toContain("240 rows");
    expect(text).toContain("First, Last");
  });

  it("says a header row with no data under it is not data", () => {
    expect(paneFor({ ...ready, paste: "First\tLast", rows: 0 }, "fields").textContent).toContain("header");
  });

  it("writes a pasted cell as TEXT, never as markup", () => {
    // Same rule as the field chips: this is a file somebody pasted, and
    // innerHTML here runs a script tag with the add-in's own privileges.
    const nasty = "<img src=x onerror=alert(1)>";
    const pane = paneFor({ ...ready, paste: `Name\n${nasty}`, columns: [nasty], rows: 1 }, "fields");
    expect(pane.querySelectorAll("img")).toHaveLength(0);
    expect((pane.querySelector("textarea") as HTMLTextAreaElement).value).toContain(nasty);
  });
});

describe("going back", () => {
  it("offers the previous step on every screen but the first", () => {
    expect(paneFor(ready, "template").querySelector("[data-back]")).toBeNull();
    for (const step of ["fields", "preview", "merge"] as const) {
      expect(paneFor(ready, step).querySelector("[data-back]"), step).not.toBeNull();
    }
  });

  it("names the step it goes to", () => {
    expect(paneFor(ready, "merge").querySelector("[data-back]")?.getAttribute("data-back")).toBe("preview");
  });

  it("never takes the last place on the screen from the primary", () => {
    // ONE PRIMARY per screen, always last. A back link that lands after it
    // makes two things to press and neither is obviously the one.
    for (const step of ["fields", "preview", "merge"] as const) {
      const main = paneFor(ready, step).querySelector("main");
      expect(main?.lastElementChild?.className, step).toBe("primary");
    }
  });
});

describe("what the host said", () => {
  it("is shown, and kept apart from what the step needs", () => {
    // "PowerPoint would not name every slide between 4 and 6" is not something
    // the user did wrong, and filing it under the same sentence as "Attach
    // your data first" makes both read as nagging.
    const pane = paneFor({ ...ready, notice: "PowerPoint would not name every slide." }, "merge");
    expect(pane.querySelector(".notice")?.textContent).toContain("PowerPoint");
  });
});

describe("the preview step", () => {
  it("says the preview is not built rather than offering it, in the HEADING too", () => {
    // The heading was the last thing on this screen still promising a preview
    // after the button stopped. A screen whose heading and whose button
    // disagree is one the user trusts neither half of.
    const pane = paneFor(ready, "preview");
    expect(pane.querySelector("h1")?.textContent).toBe("Preview is not built yet");
    expect(pane.textContent).toContain("not done");
    expect(pane.querySelector("button.primary")?.textContent).toBe("Continue to merge");
  });

  it("still carries the undo card while a preview IS showing", () => {
    const pane = paneFor({ ...ready, previewing: true }, "preview");
    expect(pane.querySelectorAll(".card.undo")).toHaveLength(1);
    expect(pane.querySelector("h1")?.textContent).toBe("A row is on the slide");
  });
});
