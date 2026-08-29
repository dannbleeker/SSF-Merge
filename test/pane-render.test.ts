// @vitest-environment jsdom
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { render } from "../src/pane/render.js";
import type { PaneState } from "../src/pane/steps.js";
import { readPastedTable, STEPS } from "../src/pane/steps.js";
import { toRecordSet } from "../src/core/data/recordset.js";

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
    expect(paneFor(ready, "fields").querySelector(".step-of")?.textContent).toBe("Step 3 of 5 · Fields");
  });

  it("draws one rail segment per step, whatever the step", () => {
    expect(paneFor(ready, "template").querySelectorAll(".rail li")).toHaveLength(5);
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
    // The zero case names the JOB rather than the absence: the step now hands
    // the user a button per column, so it is no longer a dead end.
    expect(paneFor({ ...ready, fields: [] }, "fields").querySelector("h1")?.textContent).toBe(
      "Put your fields on the slides",
    );
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
  it("is on the data step, named for main.ts to read", () => {
    expect(paneFor(ready, "data").querySelector('textarea[data-field="paste"]')).not.toBeNull();
    // And nowhere else. It moved off the fields step when the order changed,
    // and two paste boxes on two screens are two states that can disagree.
    expect(paneFor(ready, "fields").querySelector('textarea[data-field="paste"]')).toBeNull();
  });

  it("holds what was pasted as its VALUE, so a re-render does not empty it", () => {
    const pane = paneFor({ ...ready, paste: "First\tLast\nAda\tLovelace" }, "data");
    const box = pane.querySelector("textarea");
    expect((box as HTMLTextAreaElement).value).toContain("Ada");
  });

  it("names the COLUMNS it found, not just a row count", () => {
    // A paste that came through as plain text parses into one column, and a
    // row count alone looks perfectly healthy when that happens.
    const text = paneFor({ ...ready, columns: ["First", "Last"], rows: 240 }, "data").textContent ?? "";
    expect(text).toContain("240 rows");
    expect(text).toContain("First, Last");
  });

  it("says a header row with no data under it is not data", () => {
    expect(paneFor({ ...ready, paste: "First\tLast", rows: 0 }, "data").textContent).toContain("header");
  });

  it("writes a pasted cell as TEXT, never as markup", () => {
    // Same rule as the field chips: this is a file somebody pasted, and
    // innerHTML here runs a script tag with the add-in's own privileges.
    const nasty = "<img src=x onerror=alert(1)>";
    const pane = paneFor({ ...ready, paste: `Name\n${nasty}`, columns: [nasty], rows: 1 }, "data");
    expect(pane.querySelectorAll("img")).toHaveLength(0);
    expect((pane.querySelector("textarea") as HTMLTextAreaElement).value).toContain(nasty);
  });
});

describe("going back", () => {
  it("offers the previous step on every screen but the first", () => {
    expect(paneFor(ready, "template").querySelector("[data-back]")).toBeNull();
    for (const step of ["data", "fields", "preview", "merge"] as const) {
      expect(paneFor(ready, step).querySelector("[data-back]"), step).not.toBeNull();
    }
  });

  it("names the step it goes to", () => {
    expect(paneFor(ready, "merge").querySelector("[data-back]")?.getAttribute("data-back")).toBe("preview");
  });

  it("never takes the last place on the screen from the primary", () => {
    // ONE PRIMARY per screen, always last. A back link that lands after it
    // makes two things to press and neither is obviously the one.
    for (const step of ["data", "fields", "preview", "merge"] as const) {
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

  it("still offers that way past when the merge is blocked, so step 4 is never a dead end", () => {
    // This asserted the opposite until 2026-08-28: the link was hidden whenever
    // the merge was unreachable, "because it is unreachable anyway". That turned
    // a placeholder with no column into a trap. This step's primary shows a row
    // rather than advancing, so with the link gone step 4 offered "Preview the
    // first row" and "Back to fields" and nothing else, forever.
    //
    // The test kit's own template does exactly that with `{{Nickname}}`, and a
    // real run against PowerPoint for the web could not reach step 5 at all.
    //
    // Walking onto a blocked step is how the user is TOLD. It names the reason
    // and keeps its own way back, so nothing is lost by letting them arrive.
    // Held against a state that STILL blocks the merge. `{{Nickname}}` was the
    // original one and stopped being a blocker on 2026-08-29, when a
    // placeholder with no column became a caution — so keeping it here would
    // have left this test passing over a step nothing blocks, which is not the
    // thing it is for.
    const blocked: PaneState = { ...ready, fields: [] };
    expect(paneFor(blocked, "preview").querySelector("[data-forward]")?.getAttribute("data-forward")).toBe("merge");

    const merge = paneFor(blocked, "merge");
    expect(merge.querySelector(".blocked")?.textContent, "a blocked step must say why").toBeTruthy();
    expect(merge.querySelector("[data-back]"), "a blocked step must keep its way back").not.toBeNull();
    expect((merge.querySelector("button.primary") as HTMLButtonElement).disabled).toBe(true);
  });

  it("warns about a placeholder with no column, and merges anyway", () => {
    // The gate that used to stop this was the reason step 4 could dead-end, and
    // it disagreed with the engine, with the preview step and with the manual.
    // What is left is the sentence, above a button that works.
    const missing: PaneState = { ...ready, fields: ["First", "Nickname"] };
    const merge = paneFor(missing, "merge");
    expect(merge.querySelector(".caution")?.textContent).toContain("No column for Nickname");
    expect(merge.querySelector(".caution")?.textContent).toContain("stay on the slides as written");
    expect((merge.querySelector("button.primary") as HTMLButtonElement).disabled, "the warning is a wall again").toBe(
      false,
    );
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

describe("the pane says what it is waiting on", () => {
  it("names the host call in flight", () => {
    // A merge is silent for up to two and a half minutes. A frozen "Merging…"
    // is indistinguishable from a wedged pane, and the call name is the one
    // thing a user can report that says WHERE it stopped.
    const doc = paneFor({ ...ready, running: "merge", inFlight: "inserting the merged deck" }, "merge");
    expect(doc.querySelector(".inflight")?.textContent).toContain("inserting the merged deck");
  });

  it("says nothing when no call is in flight", () => {
    expect(paneFor({ ...ready, running: "merge" }, "merge").querySelector(".inflight")).toBeNull();
  });

  it("does not leave the waiting line up after the run", () => {
    // `running` is undefined once the run ends; the sentence saying what
    // happened is the notice's job from then on.
    const doc = paneFor({ ...ready, inFlight: "inserting the merged deck" }, "merge");
    expect(doc.querySelector(".inflight")).toBeNull();
  });
});

describe("the run record is reachable", () => {
  it("shows the log once the run has finished, between markers", () => {
    // The extraction channel. A task pane is a nested cross-origin iframe with
    // no devtools a user can open, and blob downloads from one are blocked
    // (office-js#1511) — so the only way this record reaches anybody is by
    // being on screen to select and copy.
    const doc = paneFor({ ...ready, log: "  0.1s  host  issued  call=x" }, "merge");
    const pre = doc.querySelector(".runlog pre");
    expect(pre?.textContent).toContain("=== SSF MERGE RUN LOG ===");
    expect(pre?.textContent).toContain("call=x");
    expect(pre?.textContent).toContain("=== END ===");
  });

  it("shows it WHILE the run is still going, which is the case it exists for", () => {
    /**
     * REVERSES a decision this test used to record. It read "keeps it out of
     * the way while the run is still going" and asserted the record was absent
     * — on the reasoning that it would compete with the sentence saying what
     * happened.
     *
     * There is no such sentence while a run is out: the notice is cleared and
     * the only line on screen names the call being waited on. What the gate
     * did cost is the case the record exists for. A host that wedges never
     * reaches the `finally` that wrote the log, so the pane sat on "Waiting on
     * PowerPoint…" forever with nothing to copy — the run nobody can explain
     * was the one with nothing to explain it with.
     */
    const doc = paneFor({ ...ready, running: "merge", log: "  0.1s  host  issued  call=x" }, "merge");
    expect(doc.querySelector(".runlog pre")?.textContent).toContain("call=x");
  });

  it("says whether the run is over, rather than claiming it is", () => {
    // "What this run did" on a run still going is a small lie, and the
    // expensive kind: it invites a reader to take the last line as the final
    // one and conclude the run stopped there.
    const going = paneFor({ ...ready, running: "merge", log: "a line" }, "merge");
    expect(going.querySelector(".runlog summary")?.textContent).toBe("What this run has done so far");
    const done = paneFor({ ...ready, log: "a line" }, "merge");
    expect(done.querySelector(".runlog summary")?.textContent).toBe("What this run did, step by step");
  });

  it("is collapsed, so a 500-line log does not bury the outcome", () => {
    const doc = paneFor({ ...ready, log: "a line" }, "merge");
    const details = doc.querySelector(".runlog");
    expect(details?.tagName.toLowerCase()).toBe("details");
    expect(details?.hasAttribute("open")).toBe(false);
  });
});

describe("a finished merge offers the slides back", () => {
  // `deckAtStart` belongs with `added`: the card's range comes from `sweepPlan`
  // now, and a positional offer needs both — how many slides, and from where.
  // Without it there is no offer at all, which is the right answer rather than
  // a guessed range.
  const landed = { ...ready, added: 720, deckSize: 732, deckAtStart: 12 };

  it("names the slides it would remove, not just 'undo'", () => {
    // This deletes part of somebody's presentation. The sentence says which
    // part, so a user who has scrolled away can check before pressing.
    const card = paneFor(landed, "merge").querySelector(".card.undo");
    expect(card?.textContent).toContain("Remove slides 13 to 732");
  });

  it("offers it only where the merge happened", () => {
    // The card belongs to the merge step. On the earlier steps the user is
    // still setting the run up, and a way to undo it there is noise.
    for (const step of ["template", "fields", "preview"] as const) {
      expect(paneFor(landed, step).querySelector(".card.undo"), step).toBeNull();
    }
  });

  it("is not offered when nothing was added", () => {
    expect(paneFor({ ...ready }, "merge").querySelector(".card.undo")).toBeNull();
    expect(paneFor({ ...ready, added: 0, deckAtStart: 12 }, "merge").querySelector(".card.undo")).toBeNull();
  });

  it("is not offered when the deck has moved on since the merge", () => {
    // Somebody appended five slides after the run. The last 720 are no longer
    // the run's own, `sweepPlan` refuses, and the card has to refuse with it —
    // it used to name a range counted back from the end of the deck, which in
    // this case is 725 of somebody else's slides.
    expect(paneFor({ ...landed, deckSize: 737 }, "merge").querySelector(".card.undo")).toBeNull();
  });

  it("disables the button while a call is out", () => {
    // The same rule the merge button follows, and for the same reason: a
    // second press during a sweep would ask the host to delete twice.
    const button = paneFor({ ...landed, running: "undo" }, "merge").querySelector(".card.undo button");
    expect(button?.hasAttribute("disabled")).toBe(true);
    expect(button?.textContent).toContain("Removing…");
  });

  it("takes the orange from the tick rather than adding a second one", () => {
    const pane = paneFor(landed, "merge");
    expect(pane.querySelector(".card.undo")).not.toBeNull();
    expect(pane.querySelector(".tick")).toBeNull();
  });
});

describe("the Insert buttons", () => {
  /**
   * The question the first person to use this add-in asked, in these words:
   * "how do I insert the fields?" The answer used to be "type them", which is
   * fine for one field and bad for twelve — and it asks the user to spell a
   * column name exactly right, from memory, with the data in another window.
   *
   * The list is the COLUMNS, in the order the data has them, because that is
   * the set the merge can fill.
   */
  const withData = { ...ready, fields: [] as string[], columns: ["First", "City"], rows: 2 };

  it("offers one per column, showing the token that will land on the slide", () => {
    const chips = Array.from(paneFor(withData, "fields").querySelectorAll("[data-insert]"));
    expect(chips.map((c) => c.getAttribute("data-insert"))).toEqual(["First", "City"]);
    expect(chips.map((c) => c.textContent)).toEqual(["{{First}}", "{{City}}"]);
  });

  it("marks a column already on the slides rather than dropping it", () => {
    // A template can legitimately use a column twice, and a list that shrinks
    // as it is pressed loses its own order under the user's hand.
    const pane = paneFor({ ...withData, fields: ["First"] }, "fields");
    expect(pane.querySelector('[data-insert="First"]')?.getAttribute("data-placed")).toBe("yes");
    expect(pane.querySelector('[data-insert="City"]')?.getAttribute("data-placed")).toBe("no");
  });

  it("says where to click first, because nothing reports where the cursor is", () => {
    // `setSelectedDataAsync` on a slide with nothing selected refuses with a
    // message about a selection, which reads as a bug. Saying it ahead of the
    // press is the only version the user can act on.
    expect(paneFor(withData, "fields").textContent).toContain("Click into a text box");
  });

  it("names the columns that are not on a slide yet", () => {
    const text = paneFor({ ...withData, fields: ["First"] }, "fields").textContent ?? "";
    expect(text).toContain("Not on a slide yet: City");
  });

  it("does not offer a column the engine could not read back", () => {
    /**
     * The defect this closes shipped: the button built `{{Column}}` and the
     * engine read it with `FIELD`, and nothing checked that those two agree.
     *
     * The dangerous case is not a header that fails to match — it is one that
     * matches a DIFFERENT, shorter name. `Total|EUR` puts a field called
     * "Total" on the slide, bound to a column that does not exist, silently.
     */
    const pane = paneFor({ ...withData, columns: ["First", "Total|EUR"] }, "fields");
    expect(pane.querySelector('[data-insert="First"]')).not.toBeNull();
    expect(pane.querySelector('[data-insert="Total|EUR"]'), "offered a token the engine misreads").toBeNull();
  });

  it("names the column it will not offer, and why", () => {
    // Named rather than silently dropped: the fix is to rename the column, and
    // a chip that is simply absent says nothing about which one or why.
    const text = paneFor({ ...withData, columns: ["First", "Total|EUR"] }, "fields").textContent ?? "";
    expect(text).toContain("Total|EUR cannot be a field");
    expect(text).toContain("Rename the column");
  });

  it("offers the headers an Excel pivot table produces", () => {
    // Reported from a real run. `Row Labels` and `Min. of cost` are the literal
    // defaults, and the reader's character class had no space on it, so the
    // pane put tokens on the slide that it then could not see.
    const columns = ["Row Labels", "Min. of cost", "Sum of quantity monthly"];
    const chips = Array.from(paneFor({ ...withData, columns }, "fields").querySelectorAll("[data-insert]"));
    expect(chips.map((c) => c.getAttribute("data-insert"))).toEqual(columns);
  });

  it("withholds the stale empty-slides line while a note is up", () => {
    /**
     * An insert lands on the SLIDE and tells the pane nothing, so between a
     * fresh note and a sentence read off the last template read, the note is
     * the current one. Without this the screen said "{{City}} put on the
     * slide" directly above "these slides carry no fields yet" — contradicting
     * itself about something the user had just done.
     */
    const pane = paneFor({ ...withData, fields: [], fieldNote: "{{City}} put on the slide." }, "fields");
    expect(pane.textContent).toContain("put on the slide");
    expect(pane.textContent, "said both at once").not.toContain("carry no fields yet");
    // And it is back the moment there is no note to contradict it.
    expect(paneFor({ ...withData, fields: [] }, "fields").textContent).toContain("carry no fields yet");
  });

  it("says nothing about unused columns before anything is placed", () => {
    // Before the first insert every column is unused and the line says nothing.
    expect(paneFor(withData, "fields").textContent).not.toContain("Not on a slide yet");
  });

  it("writes a column name as TEXT, never as markup", () => {
    // Same rule as the field chips and the paste box: this name came out of a
    // file somebody pasted, and innerHTML here runs a script tag with the
    // add-in's own privileges.
    const nasty = "<img src=x onerror=alert(1)>";
    const pane = paneFor({ ...withData, columns: [nasty] }, "fields");
    expect(pane.querySelectorAll("img")).toHaveLength(0);
    expect(pane.querySelector("[data-insert]")?.textContent).toBe(`{{${nasty}}}`);
  });

  it("says what the last press did, and leaves the primary last", () => {
    const pane = paneFor({ ...withData, fieldNote: "{{City}} put on the slide." }, "fields");
    expect(pane.textContent).toContain("{{City}} put on the slide.");
    const controls = Array.from(pane.querySelectorAll("button"));
    expect(controls[controls.length - 1]?.className).toContain("primary");
  });

  it("shows a chart's placeholder as an ordinary chip", () => {
    // It used to have a card of its own saying it could not be filled. Now the
    // merge fills it, so it is a field like any other and the step that lists
    // what is on the slides lists it.
    const pane = paneFor({ ...withData, fields: ["First", "Region"] }, "fields");
    expect(Array.from(pane.querySelectorAll(".fields li")).map((n) => n.textContent)).toEqual(["First", "Region"]);
    expect(pane.textContent, "the stale refusal").not.toContain("SmartArt");
  });

  it("says why the list is empty rather than showing nothing", () => {
    // Reachable through the Back link from Preview after the paste box is
    // cleared. An empty list on a step whose whole job is inserting reads as
    // the control being broken.
    const pane = paneFor({ ...ready, fields: [], columns: undefined, rows: undefined }, "fields");
    expect(pane.textContent).toContain("Attach your data first");
    expect(pane.querySelector("[data-insert]")).toBeNull();
  });
});

describe("the picture picker", () => {
  const paste = "Name,Photo\nAda,ada.png\nGrace,grace.jpg";
  const read = readPastedTable(paste);
  const withPhotos: PaneState = {
    ...ready,
    paste,
    records: read.records ?? undefined,
    columns: read.columns,
    rows: read.rows,
  };
  const files = (...names: string[]) => new Map(names.map((n) => [n, new Uint8Array([1])]));

  it("appears on the data step, where the rest of the data is collected", () => {
    expect(paneFor(withPhotos, "data").querySelector('[data-field="images"]')).not.toBeNull();
  });

  it("stays away when no column names a picture", () => {
    const plain = "Name,City\nAda,London";
    const flat = readPastedTable(plain);
    const state: PaneState = {
      ...ready,
      paste: plain,
      records: flat.records ?? undefined,
      columns: flat.columns,
      rows: flat.rows,
    };
    expect(paneFor(state, "data").querySelector('[data-field="images"]')).toBeNull();
  });

  it("says how many pictures the data asks for, and which column asked", () => {
    expect(paneFor(withPhotos, "data").textContent).toContain("2 pictures named in Photo");
  });

  it("says what skipping costs, because skipping is allowed", () => {
    expect(paneFor(withPhotos, "data").textContent).toContain("keep the placeholder");
  });

  it("reports every picture matched once they are all picked", () => {
    const pane = paneFor({ ...withPhotos, images: files("ada.png", "grace.jpg") }, "data");
    expect(pane.textContent).toContain("All 2 pictures matched.");
    expect(pane.querySelector(".images .blocked")).toBeNull();
  });

  it("names the missing ones and marks the line, rather than counting them", () => {
    const pane = paneFor({ ...withPhotos, images: files("ada.png") }, "data");
    expect(pane.querySelector(".images .blocked")?.textContent).toBe("1 of 2 matched. Missing: grace.jpg.");
  });

  it("mentions files no row refers to once, and not as a problem", () => {
    const pane = paneFor({ ...withPhotos, images: files("ada.png", "grace.jpg", "logo.png") }, "data");
    expect(pane.textContent).toContain("1 file no row refers to — ignored.");
    expect(pane.querySelector(".images .blocked")).toBeNull();
  });

  it("takes only the picture types the engine can read", () => {
    const input = paneFor(withPhotos, "data").querySelector('[data-field="images"]');
    expect(input?.getAttribute("accept")).toBe("image/png,image/jpeg,image/gif,image/bmp");
    expect(input?.hasAttribute("multiple")).toBe(true);
  });
});

describe("the sentence above the merge button", () => {
  /**
   * It promised nine slides where the plan built eight, because the count was
   * slides-per-record times rows and a conditional slide is not produced for
   * every row.
   *
   * Asserted HERE and not only on `plannedSlides`, because the unit test passes
   * happily while the card goes on calling the old product — which is what a
   * mutation showed: reverting the call site broke nothing until this existed.
   */
  const records = toRecordSet([
    ["First", "Renewal"],
    ["Ada", "yes"],
    ["Bo", "no"],
    ["Cy", "yes"],
    ["Di", "no"],
  ]);

  const withCondition: PaneState = {
    block: { from: 3, to: 5 },
    fields: ["First"],
    columns: ["First", "Renewal"],
    records,
    rows: 4,
    previewing: false,
    deckSize: 10,
    // The middle slide only for a renewal, and one row taken out.
    conditions: { 4: "Renewal" },
    excluded: [3],
  };

  it("counts the conditional slides out", () => {
    const text = paneFor(withCondition, "merge").querySelector(".facts")?.textContent ?? "";
    // Three rows merge; the middle slide is produced for one of them.
    expect(text).toContain("8 slides added after slide 10");
    expect(text).toContain("18 slides in the deck");
    // The product the card used to show, which was over by one.
    expect(text, "the card is showing rows x slides again").not.toContain("9 slides added");
  });

  it("still shows the plain product when nothing is conditional", () => {
    const plain = { ...withCondition, conditions: undefined };
    const text = paneFor(plain, "merge").querySelector(".facts")?.textContent ?? "";
    expect(text).toContain("9 slides added after slide 10");
  });
});

describe("every control the pane draws can be named", () => {
  /**
   * A control with no accessible name is announced as its type and nothing
   * else — "button", "edit" — so a screen-reader user is told there is
   * something here and not what it does.
   *
   * Swept across every step rather than asserted control by control, because
   * the one that was missing had been added later than the pattern: every other
   * control on this pane sits inside a `<label>` carrying its caption, and the
   * picture picker is a `<div>` — it holds a tally and a missing-file list as
   * well — so its caption sat beside the input and was attached to nothing.
   * It is the only control on the pane that attaches the pictures.
   *
   * Three rules, all of them things a browser or a reader will act on: a name
   * for every control, no duplicated id, and no button that says nothing.
   */
  const records = toRecordSet([
    ["Name", "Photo", "Renewal"],
    ["Ada", "ada.png", "yes"],
    ["Bo", "bo.png", "no"],
  ]);

  /** Everything open at once, so every branch of the pane is drawn. */
  const rich: PaneState = {
    block: { from: 3, to: 5 },
    fields: ["Name", "Photo", "Nickname"],
    imageFields: ["Photo"],
    columns: ["Name", "Photo", "Renewal"],
    records,
    rows: 2,
    previewing: false,
    deckSize: 10,
    draft: { from: "3", to: "5" },
    paste: "Name\tPhoto\nAda\tada.png",
    conditions: { 4: "Renewal" },
    rowsOpen: true,
    conditionsOpen: true,
    notice: "something the host said",
    fieldNote: "{{Name}} put on the slide.",
  };

  it.each(STEPS.map((s) => [s]))("on step %s", (step) => {
    const root = paneFor(rich, step);
    const faults: string[] = [];

    const seen = new Map<string, number>();
    for (const node of Array.from(root.querySelectorAll("[id]"))) {
      const id = node.getAttribute("id") as string;
      seen.set(id, (seen.get(id) ?? 0) + 1);
    }
    for (const [id, n] of seen) if (n > 1) faults.push(`id "${id}" appears ${n} times`);

    for (const control of Array.from(root.querySelectorAll("input, select, textarea"))) {
      const named =
        control.closest("label") !== null ||
        control.hasAttribute("aria-label") ||
        control.hasAttribute("aria-labelledby") ||
        (control.id !== "" && root.querySelector(`label[for="${control.id}"]`) !== null);
      if (!named)
        faults.push(`<${control.tagName.toLowerCase()} ${control.getAttribute("data-field") ?? ""}> has no name`);
    }

    for (const button of Array.from(root.querySelectorAll("button"))) {
      if ((button.textContent ?? "").trim() === "" && !button.hasAttribute("aria-label")) {
        faults.push("a button says nothing");
      }
    }

    expect(faults).toEqual([]);
    // A step that drew no controls at all would pass the three rules above
    // while proving nothing, so say that it drew something.
    expect(root.querySelectorAll("button, input, select, textarea").length).toBeGreaterThan(0);
  });
});

/**
 * The one thing this suite cannot measure, asserted where it is decided.
 *
 * jsdom has no layout, so no test here can see a pane that has gone wider than
 * the iframe it lives in. What it can see is the rule that stops it, and this
 * is the same shape as the source scans elsewhere in the repo: weaker than a
 * measurement, and the only guard available for a defect that was found by
 * measuring.
 *
 * Found at 320px, the narrow end of the width `taskpane.css` designs for, with
 * a real column header. A pivot table's default headers are what this add-in
 * is pasted the most, and one of them with no spaces in it took the document
 * to 545px; a spaceless host error took it to 3751px, with the primary button
 * off the side of a pane the user cannot scroll usefully.
 */
describe("nothing a user supplies may push the pane sideways", () => {
  const css = readFileSync("src/pane/taskpane.css", "utf8");

  it("wraps inside a word, on the element everything inherits from", () => {
    const body = css.match(/\nbody \{([\s\S]*?)\n\}/)?.[1] ?? "";
    expect(body, "found the body block at all").not.toBe("");
    // `anywhere`, never `break-word`: only `anywhere` shrinks a flex or grid
    // item's min-content width, and the chips, cards and button rows are flex.
    // `break-word` leaves all three measured overflows exactly where they were.
    expect(body).toContain("overflow-wrap: anywhere");
  });

  it("never paints TEXT with the blue that is a background colour", () => {
    // Measured with the pane's own renderer in a browser, in both themes:
    // `--blue` stays dark because it is a FILL — the header and the primary
    // button, both carrying white text — and read as text on the dark
    // surface it is 3.0:1. Three declarations did that, and one of them was
    // every secondary button, "Remove these slides" among them, at 2.93:1.
    //
    // `--link` is the same colour in light and a lighter one in dark, which
    // is why the dark blocks define it at all. Anything blue that is text
    // reads it.
    const asText = [...css.matchAll(/(^|[^-])color: var\(--blue\)/g)];
    expect(asText.map((m) => m[0])).toEqual([]);
  });

  it("leaves the two places that deliberately do not wrap", () => {
    // The rule reaches text that wraps, so its stated scope is only true while
    // these two keep saying so. The run log is a scroll box a user copies out
    // of, and a row label is one line with an ellipsis at 320px.
    expect(css).toMatch(/\.runlog pre \{[^}]*white-space: pre/);
    expect(css).toMatch(/\.rowlist label \{[^}]*white-space: nowrap/);
  });
});
