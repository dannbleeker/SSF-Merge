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
      // TWO unmatched placeholders — the ordinary case of a paste missing a
      // couple of columns, and the one every fixture here happened to miss.
      { ...ready, fields: ["First", "Nickname", "Badge"] },
      { ...ready, fields: ["First", "Nickname", "Badge"], previewing: true },
      // A run in flight, and a run that landed.
      { ...ready, running: "merge" as const },
      { ...ready, running: "inspect" as const },
      { ...ready, added: 720 },
    ];
    for (const state of states) {
      for (const step of ["template", "fields", "preview", "merge"] as const) {
        const pane = paneFor(state, step);
        // HOLDERS, not elements. A row of chips is one signal — "these are the
        // ones with no column" — and `orangeHolder` already models it that
        // way, returning a single holder name. Counting chips instead said a
        // template missing two columns broke a budget nothing had broken: an
        // ordinary state, which no fixture in the sweep or in the screenshot
        // script happened to reach, so CI reported neither the violation nor
        // the miscount.
        const holders = [".tick", ".card.undo", '.fields li[data-matched="no"]'].filter(
          (sel) => pane.querySelectorAll(sel).length > 0,
        );
        expect(holders.length, `${step} with ${JSON.stringify(state)}`).toBeLessThanOrEqual(1);
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

  it("sets the box's VALUE, and leaves the value ATTRIBUTE alone", () => {
    // The `.value` half of this passed against `setAttribute("value", …)` too
    // — a fresh input reflects the content attribute into the property, and
    // `render` builds fresh elements every time — so it could not fail against
    // the implementation it is named for. The attribute assertion is the half
    // that discriminates.
    //
    // The reason is not the one first written down here either. It is not that
    // a box would "snap back": it is that a TEXTAREA has no value attribute at
    // all, so one helper serving both controls has to write the property.
    const pane = paneFor({ fields: [], previewing: false, draft: { from: "4", to: "6" } }, "template");
    const from = pane.querySelector('input[data-field="from"]');
    expect(from).toBeInstanceOf(HTMLInputElement);
    expect((from as HTMLInputElement).value).toBe("4");
    expect((from as HTMLInputElement).getAttribute("value"), "the property, not the attribute").toBeNull();
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
  it("says what pressing it will do, with the number in it", () => {
    const pane = paneFor(ready, "preview");
    expect(pane.querySelector("h1")?.textContent).toBe("See one row before you commit");
    expect(pane.textContent).toContain("Adds 3 slides");
    expect(pane.querySelector("button.primary")?.textContent).toBe("Preview the first row");
  });

  it("offers a way past it, because its primary does not advance", () => {
    // Every other step's primary carries the user forward; this one shows a
    // row. Without a forward link the wizard has no exit from step 3 at all,
    // which is exactly what making the preview real took away.
    const pane = paneFor(ready, "preview");
    expect(pane.querySelector("[data-forward]")?.getAttribute("data-forward")).toBe("merge");
  });

  it("does not offer that way past when the merge is not reachable anyway", () => {
    const missing: PaneState = { ...ready, fields: ["First", "Nickname"] };
    expect(paneFor(missing, "preview").querySelector("[data-forward]")).toBeNull();
  });

  it("names the slides a preview landed on, so a closed pane is recoverable", () => {
    // A user who closes the pane mid-preview has no other way to find out
    // which slides to delete.
    const pane = paneFor({ ...ready, previewing: true, previewSlides: { from: 13, to: 15 } }, "preview");
    expect(pane.querySelectorAll(".card.undo")).toHaveLength(1);
    expect(pane.querySelector(".card.undo")?.textContent).toContain("Slides 13 to 15");
    expect(pane.querySelector("h1")?.textContent).toBe("The first row is in your deck");
  });

  it("says the template is untouched, which is the obvious fear", () => {
    const card = paneFor({ ...ready, previewing: true, previewSlides: { from: 13, to: 13 } }, "preview").querySelector(
      ".card.undo",
    );
    expect(card?.textContent).toContain("template is untouched");
    // Singular, because one slide is one slide.
    expect(card?.textContent).toContain("Slide 13 ");
  });

  it("hides the forward link while a preview is showing", () => {
    // It would carry the user to a merge step that refuses to merge.
    const pane = paneFor({ ...ready, previewing: true, previewSlides: { from: 13, to: 15 } }, "preview");
    expect(pane.querySelector("[data-forward]")).toBeNull();
  });
});

describe("the merge card is a forecast", () => {
  it("goes as soon as the merge starts, because it reads as already done", () => {
    // "720 slides added after slide 12" over a button reading "Merging…"
    // announces slides that are not in the deck yet. Found by looking at the
    // screenshot, which is the only thing that could see it.
    const running = paneFor({ ...ready, running: "merge" }, "merge");
    expect(running.querySelectorAll(".card.summary")).toHaveLength(0);
    expect(running.querySelector("button.primary")?.textContent).toBe("Merging…");
  });

  it("does not sit beside the notice saying the same thing afterwards", () => {
    const done = paneFor({ ...ready, added: 720, notice: "720 slides added after slide 12." }, "merge");
    expect(done.querySelectorAll(".card.summary")).toHaveLength(0);
    expect(done.querySelector(".notice")?.textContent).toContain("720 slides added");
  });

  it("is still there before anything has been pressed", () => {
    expect(paneFor(ready, "merge").querySelectorAll(".card.summary")).toHaveLength(1);
  });
});

describe("the row picker", () => {
  const records = {
    columns: [{ name: "Name", type: "text" as const }],
    rows: [{ Name: "Ada" }, { Name: "Grace" }, { Name: "Katherine" }],
  };
  const withData: PaneState = { ...ready, records, rows: 3 };

  it("is CLOSED by default — 240 rows is not a screen", () => {
    const pane = paneFor(withData, "merge");
    expect(pane.querySelector('[data-action="rows"]')).not.toBeNull();
    expect(pane.querySelectorAll(".rowlist")).toHaveLength(0);
  });

  it("says how many rows there are while it is shut", () => {
    expect(paneFor(withData, "merge").querySelector('[data-action="rows"]')?.textContent).toContain("All 3 rows");
  });

  it("says how many were taken out, once some were", () => {
    const some = { ...withData, excluded: [1] };
    expect(paneFor(some, "merge").querySelector('[data-action="rows"]')?.textContent).toContain("1 row taken out");
  });

  it("lists the rows with a checkbox each when open", () => {
    const pane = paneFor({ ...withData, rowsOpen: true }, "merge");
    const boxes = pane.querySelectorAll<HTMLInputElement>('.rowlist input[type="checkbox"]');
    expect(boxes).toHaveLength(3);
    expect(Array.from(boxes).every((b) => b.checked)).toBe(true);
  });

  it("unticks the rows that are out, as a PROPERTY not an attribute", () => {
    // Same reason the slide-number boxes use the property: `checked` as an
    // attribute is the default a control reverts to, and this pane rebuilds
    // itself on every change.
    const pane = paneFor({ ...withData, rowsOpen: true, excluded: [1] }, "merge");
    const boxes = pane.querySelectorAll<HTMLInputElement>('.rowlist input[type="checkbox"]');
    expect(Array.from(boxes).map((b) => b.checked)).toEqual([true, false, true]);
    expect(boxes[1]?.hasAttribute("checked")).toBe(false);
  });

  it("filters the list by the search box", () => {
    const pane = paneFor({ ...withData, rowsOpen: true, rowSearch: "grace" }, "merge");
    expect(pane.querySelectorAll(".rowlist li")).toHaveLength(1);
    expect(pane.querySelector(".rowlist li")?.textContent).toContain("Grace");
  });

  it("says so when nothing matches, rather than showing an empty box", () => {
    const pane = paneFor({ ...withData, rowsOpen: true, rowSearch: "zzz" }, "merge");
    expect(pane.textContent).toContain("No row matches that");
  });

  it("COUNTS what it did not show rather than dropping it silently", () => {
    // A list that stops without saying so is one the user believes they have
    // read. 80 rows against a cap of 60.
    const many = {
      columns: records.columns,
      rows: Array.from({ length: 80 }, (_, i) => ({ Name: `Person ${i}` })),
    };
    const pane = paneFor({ ...ready, records: many, rows: 80, rowsOpen: true }, "merge");
    expect(pane.querySelectorAll(".rowlist li")).toHaveLength(60);
    expect(pane.textContent).toContain("80 matches");
  });

  it("puts the INCLUDED count in the summary card, not the pasted one", () => {
    const some = { ...withData, excluded: [1, 2], deckSize: 12 };
    // 1 row x 3 slides.
    expect(paneFor(some, "merge").querySelector(".facts")?.textContent).toContain("3 slides added");
  });

  it("writes a row's label as TEXT — it came from a pasted file", () => {
    const nasty = "<img src=x onerror=alert(1)>";
    const pane = paneFor(
      { ...ready, records: { columns: records.columns, rows: [{ Name: nasty }] }, rows: 1, rowsOpen: true },
      "merge",
    );
    expect(pane.querySelectorAll("img")).toHaveLength(0);
    expect(pane.querySelector(".rowlist li")?.textContent).toContain(nasty);
  });
});
