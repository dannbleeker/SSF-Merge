import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import { withoutTsComments, withoutTsProse } from "../scripts/without-prose.mjs";
import { creationIdOf } from "../src/core/pptx/clone.js";
import { Pkg } from "../src/core/pptx/pkg.js";
import { TAG_RUN, readSlideTags } from "../src/core/pptx/tags.js";
import { Q3, Q4 } from "../src/host/verdicts.js";

/**
 * These read the ARTIFACT, not a recipe for it.
 *
 * The probe's whole value is that its two arms differ in exactly one way, and a
 * builder that quietly produced two identical decks would give an answer sheet
 * that looks complete and says nothing. Reading the generated snippet means
 * there is no second copy of the recipe to drift from the first, and it fails
 * when the snippet has not been regenerated.
 */
const snippet = readFileSync("probe/probe-snippet.ts", "utf8");

function deckFromSnippet(name: string): string {
  // Tolerant of whitespace on purpose: a formatter that wrapped the
  // assignment once broke this silently, and a guard that stops matching is a
  // guard that passes forever.
  const match = new RegExp(`const ${name}\\s*=\\s*"([A-Za-z0-9+/=]+)"`).exec(snippet);
  if (!match?.[1]) throw new Error(`the snippet has no ${name}`);
  return match[1];
}

describe("the probe's fixture decks", () => {
  it("both hold two slides, so the arms differ only in the creation ids", async () => {
    for (const name of ["FRESH_DECK", "COLLISION_DECK"]) {
      const pkg = await Pkg.open(deckFromSnippet(name));
      expect(await pkg.slidePaths(), name).toHaveLength(2);
    }
  });

  it("gives the fresh arm two DIFFERENT creation ids", async () => {
    const pkg = await Pkg.open(deckFromSnippet("FRESH_DECK"));
    const ids = await Promise.all((await pkg.slidePaths()).map((p) => creationIdOf(pkg, p)));
    expect(new Set(ids).size).toBe(2);
  });

  it("gives the collision arm ONE creation id shared by both slides", async () => {
    // The arm that tests office-js#6105. If this ever produced two distinct ids
    // the probe would report a clean insert and prove nothing, which is the
    // vacuous measurement this project keeps having to catch.
    const pkg = await Pkg.open(deckFromSnippet("COLLISION_DECK"));
    const ids = await Promise.all((await pkg.slidePaths()).map((p) => creationIdOf(pkg, p)));
    expect(new Set(ids).size).toBe(1);
    expect(ids[0]).toBeDefined();
  });

  it("tags the fresh arm's slides in the package, which is what question two reads back", async () => {
    const pkg = await Pkg.open(deckFromSnippet("FRESH_DECK"));
    for (const path of await pkg.slidePaths()) {
      expect((await readSlideTags(pkg, path)).get(TAG_RUN), path).toBe("probe-run");
    }
  });
});

describe("the probe snippet", () => {
  it("bounds every host call, because a stall here is death rather than slowness", () => {
    expect(snippet).toContain("withTimeout");
    expect(snippet).not.toMatch(/await PowerPoint\.run\([\s\S]{0,40}\)\s*;\s*\/\/ unbounded/);
  });

  it("cleans up by position and never by id", () => {
    // A slide this run added does not round-trip through slides.getItem(id) on
    // the web, and a sibling project's by-id clean-up reported 45 successful
    // deletes having removed nothing.
    expect(snippet).toContain("getItemAt(i).delete()");
    expect(snippet).not.toContain("getItem(id).delete");
  });

  it("passes a targetSlideId, so the probe's slides do not land in front of the user's", () => {
    expect(snippet).toContain("targetSlideId");
  });

  it("carries a floor that stops the sweep reaching the user's own slides", () => {
    expect(snippet).toContain("from < deckAtStart");
  });
});

describe("the probe's fixture package", () => {
  it("carries a theme with all three of its required children", async () => {
    // CT_BaseStyles requires clrScheme, fontScheme and fmtScheme, every one
    // mandatory. This part was `<a:themeElements/>` when the first real sheet
    // came back InvalidArgument from every insert, and KeepSourceFormatting is
    // exactly the path that has to import the source theme. Checked against a
    // deck PowerPoint itself accepts, which carries all three.
    const pkg = await Pkg.open(deckFromSnippet("FRESH_DECK"));
    const theme = await pkg.text("ppt/theme/theme1.xml");
    for (const child of ["clrScheme", "fontScheme", "fmtScheme"]) {
      expect(theme).toContain(`<a:${child} `);
    }
  });

  it("names every part it declares a content type for", async () => {
    // A content-type Override for a part that is not in the zip is a package
    // no reader will open, and it is the easy mistake when adding parts.
    const pkg = await Pkg.open(deckFromSnippet("FRESH_DECK"));
    const types = await pkg.text("[Content_Types].xml");
    const parts = [...types.matchAll(/<Override PartName="\/([^"]+)"/g)].map((m) => m[1] ?? "");
    expect(parts.length).toBeGreaterThan(4);
    for (const part of parts) {
      await expect(pkg.text(part)).resolves.toBeTypeOf("string");
    }
  });
});

describe("the generated snippet", () => {
  it("typechecks against the real Office.js types", () => {
    // The snippet is outside tsconfig's include and is pasted into an editor
    // that will run it before anyone reads it, so nothing else here would catch
    // a misspelled option key or a call that does not exist. Proven non-vacuous
    // by adding an unknown key to InsertSlideOptions: tsc names it.
    expect(() =>
      execFileSync(
        process.execPath,
        [
          "./node_modules/typescript/bin/tsc",
          "--noEmit",
          "--lib",
          "es2020,dom",
          "--types",
          "office-js",
          "--skipLibCheck",
          "probe/probe-snippet.ts",
        ],
        { encoding: "utf8", stdio: "pipe" },
      ),
    ).not.toThrow();
  }, 60000);

  it("asks its control arm BEFORE it adds anything of its own", () => {
    // The control inserts the presentation's own bytes. Run later it would be
    // inserting the probe's slides back too, which grows the deck by whatever
    // the earlier arms happened to land and makes its own reading unreadable.
    const control = snippet.indexOf("answers.insertSelf");
    const first = snippet.indexOf("answers.insertFresh");
    expect(control).toBeGreaterThan(-1);
    expect(control).toBeLessThan(first);
  });

  it("names the call that threw rather than the whole probe", () => {
    // Four calls shared one catch, so the first real sheet said
    // "InvalidArgument" about a statement nobody could identify.
    expect(snippet).toContain("failedAt: step");
    // Every step the probe can be inside names the batch it is in, so the sheet
    // says which one raised. The second real sheet named its own failure this
    // way on the first outing.
    const steps = [...snippet.matchAll(/step = "([^"]+)"/g)].map((m) => m[1]);
    expect(steps.length).toBeGreaterThan(2);
    expect(steps).toContain("question 3: create, style and write in one batch");
  });
});

describe("the substring experiments", () => {
  it("uses the strings the reader scores against", () => {
    // The reader once expected "Hello Ada here and 1-2".replace("2", "BBB") —
    // a string neither model produces. Both sides read one constant now, and
    // this fails when the snippet has not been regenerated after a change.
    expect(snippet).toContain(`const Q3_TEXT = "${Q3.text}"`);
    expect(snippet).toContain(`const Q4_TEXT = "${Q4.text}"`);
  });

  it("predicts two DIFFERENT strings for the two offset models", () => {
    // A guard against the pair silently collapsing: if a change made the
    // independent and shifted answers equal, question four would report "yes"
    // whatever the host did.
    expect(Q4.independent).not.toBe(Q4.shifted);
  });

  it("predicts answers both models can actually produce", () => {
    // Applying each model to the text by hand. If neither prediction is what
    // the arithmetic gives, the question is scored against fiction.
    const both = Q4.text.slice(0, 0) + "XXXXX" + Q4.text.slice(3);
    expect(both.slice(0, 4) + "2" + both.slice(7)).toBe(Q4.shifted);
    expect(Q4.text.slice(0, 0) + "XXXXX" + Q4.text.slice(3, 4) + "2").toBe(Q4.independent);
  });

  it("touches no shape proxy across a sync", () => {
    // The defect the second real sheet named: Office.js rewrites a created
    // shape's path to shapes.getItem(id) once it has been through a sync, and
    // this host answers 5010 InvalidParam for that id. Every write is queued in
    // the batch that created the shape.
    const body = snippet.slice(snippet.indexOf("async function substringProbe"), snippet.indexOf("/** Positional"));
    expect(body.match(/await context\.sync\(\)/g) ?? []).toHaveLength(2);
    expect(body).not.toContain("shape.delete()");
  });

  it("refuses to draw when the probe has no slide of its own", () => {
    // Otherwise a run whose inserts all failed leaves two text boxes on the
    // user's own slide, which is the deck it was told it could not damage.
    expect(snippet).toContain("if (!deckGrew)");
  });
});

describe("the questions the review added", () => {
  const snippet = readFileSync("probe/probe-snippet.ts", "utf8");

  it("asks whether a collection load answers short, and against what", () => {
    // office-js#4272. `getCount` is the authority — a scalar, not a load — so
    // the question is only answerable if both are read in the same breath.
    expect(snippet).toContain("deckReadProbe");
    expect(snippet).toMatch(/getCount\(\)/);
    expect(snippet).toContain("out.short");
  });

  it("asks whether a SHORT read is the first n in deck order", () => {
    // The half that decides how bad it is. A prefix-stable short read means a
    // block inside it is right and one past it is refused; a scrambled one
    // means `indexOf` answers the wrong SLIDE NUMBER and the merge clones
    // slides nobody chose. Checked against getItemAt, which is a different
    // code path from a collection load.
    expect(snippet).toContain("prefixOk");
    expect(snippet).toMatch(/getItemAt\(i\)/);
  });

  it("asks whether the read comes back EMPTY after a sync that succeeded", () => {
    // office-js#6363, and the sibling project's central failure.
    expect(snippet).toContain("out.empty");
  });

  it("says whether the deck was even big enough to answer the >50 question", () => {
    // A nine-slide deck cannot, and must not look as though it did — the
    // vacuous-measurement trap this repo keeps meeting.
    expect(snippet).toContain("canAnswerFiftyQuestion");
  });

  it("asks about inserting while a shape is selected WITHOUT selecting one", () => {
    // The workaround would be `setSelectedShapes`, which is the one call in
    // this family with a measured history of wedging the host. So the probe
    // reads what is already selected and never sets anything.
    expect(snippet).toContain("insertWhileSelectedProbe");
    expect(snippet).toContain("getSelectedShapes");
    // Against the CODE, not the prose. The snippet explains at length why it
    // does not call `setSelectedShapes`, in comments containing the name — and
    // the first version of this assertion matched those. Fourth time in this
    // repo, and the first since `without-prose.mjs` existed to prevent it: the
    // tool was there and I did not reach for it.
    expect(withoutTsProse(snippet), "never sets a selection").not.toContain("setSelectedShapes");
  });

  it("runs both new questions", () => {
    expect(snippet).toContain("answers.deckRead");
    expect(snippet).toContain("answers.insertWhileSelected");
  });
});

describe("every arm the probe collects is READ", () => {
  /**
   * The gap this closes was real and it was mine. `deckRead` and
   * `insertWhileSelected` were asked, answered by a live host, written into the
   * sheet — and `read-answers.mjs` printed neither, so the round reported seven
   * of nine arms and looked complete. An arm nobody reads is the same nothing
   * as an arm nobody asks, and it is worse, because the sheet says it is there.
   *
   * The test above this one was the near miss: it is named "reports them" and
   * asserts only that the SNIPPET runs them. It is renamed now to claim what it
   * checks.
   *
   * Both halves come from source rather than from a list here, so a list cannot
   * go stale — which is the failure mode a hand-written pair of names would
   * reintroduce on the next arm.
   */
  const reader = readFileSync("scripts/read-answers.mjs", "utf8");
  const collected = [...snippet.matchAll(/\banswers\.(\w+)\s*=/g)].map((m) => m[1]);

  it("finds arms to check", () => {
    // Guards the regex, not the reader. A pattern that matched nothing would
    // make every assertion below pass on an empty list.
    expect(collected.length).toBeGreaterThan(5);
  });

  it.each(collected)("reads %s", (arm) => {
    expect(withoutTsComments(reader), `${arm} is collected and never read`).toContain(`sheet.${arm}`);
  });
});

describe("every probe snippet runs itself", () => {
  /**
   * Script Lab's own SAMPLE snippet ends `$("#run").click(...)` — jQuery, and a
   * button that exists in the sample's HTML tab. Copying that shape into a
   * BLANK snippet throws `ReferenceError: $ is not defined` before a single
   * Office call, because a blank snippet has neither.
   *
   * `probe-snippet.ts` has always called `main()` at the top level and was
   * fine. `aspect-probe.ts` shipped with the sample's boilerplate and cost a
   * round trip to the owner's PowerPoint to find out. Both are checked here so
   * the next generated snippet cannot repeat it.
   *
   * Read off the CODE, not the file: both snippets have header comments that
   * quote the broken form on purpose, and a naive grep matches those.
   */
  const SNIPPETS = ["probe/probe-snippet.ts", "probe/aspect-probe.ts"];

  it("calls its entry point at the top level", () => {
    for (const path of SNIPPETS) {
      const code = withoutTsComments(readFileSync(path, "utf8")) as string;
      expect(code, `${path} never invokes itself`).toMatch(/^\s*(main|run)\(\)/m);
    }
  });

  it("never reaches for jQuery or a #run button a blank snippet does not have", () => {
    for (const path of SNIPPETS) {
      const code = withoutTsComments(readFileSync(path, "utf8")) as string;
      expect(code, `${path} uses jQuery`).not.toMatch(/\$\s*\(/);
      expect(code, `${path} wires a button that is not there`).not.toContain("#run");
    }
  });
});
