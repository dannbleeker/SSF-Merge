import { readdirSync, readFileSync, statSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import { withoutTsProse } from "../scripts/without-prose.mjs";

function filesUnder(dir: string): string[] {
  return readdirSync(dir).flatMap((entry) => {
    const path = join(dir, entry);
    return statSync(path).isDirectory() ? filesUnder(path) : path.endsWith(".ts") ? [path] : [];
  });
}

/**
 * One function's source, brace-matched from its signature.
 *
 * The first version sliced to the next `\n/**`, which is not where a function
 * ends — it is where the next DOCBLOCK starts, and only by accident of the
 * current file order was that the same place. An adversarial pass proved the
 * cost: with `readTemplate` reverted to the counting loop this guard exists to
 * forbid, and a helper below it holding the strings the guard greps for, the
 * slice ran past the defect into the helper and every assertion passed. With no
 * later `/**` at all, `indexOf` answers -1 and `slice(start, -1)` swallows the
 * rest of the file.
 *
 * Slicing to the next top-level `export` was tried second and is ALSO wrong: a
 * plain `function` between them is not a stop, so the same repro still passed.
 * Only the braces say where a function ends, so they are what is counted —
 * over a copy with strings and comments masked to spaces, because a brace in a
 * comment closes nothing and this file's prose is full of them.
 */
function maskLiterals(src: string): string {
  const out = src.split("");
  let i = 0;
  const blank = (from: number, to: number) => {
    for (let j = from; j < to && j < out.length; j++) if (out[j] !== "\n") out[j] = " ";
  };
  while (i < src.length) {
    const two = src.slice(i, i + 2);
    if (two === "//") {
      const end = src.indexOf("\n", i);
      blank(i, end === -1 ? src.length : end);
      i = end === -1 ? src.length : end;
    } else if (two === "/*") {
      const end = src.indexOf("*/", i + 2);
      blank(i, end === -1 ? src.length : end + 2);
      i = end === -1 ? src.length : end + 2;
    } else if (src[i] === '"' || src[i] === "'" || src[i] === "`") {
      const quote = src[i];
      let j = i + 1;
      while (j < src.length && src[j] !== quote) j += src[j] === "\\" ? 2 : 1;
      blank(i, j + 1);
      i = j + 1;
    } else i++;
  }
  return out.join("");
}

function functionBody(src: string, signature: string): string {
  const start = src.indexOf(signature);
  expect(start, `${signature} is in the file`).toBeGreaterThan(-1);
  const masked = maskLiterals(src);
  // Past the PARAMETER LIST first. The first `{` after the signature is not
  // the body when a parameter is typed inline — `readTemplate(block: { from:
  // number; to: number })` matched its own annotation, so the "body" was
  // eleven words of signature and every `toContain` failed on a file that was
  // perfectly correct. A guard that goes red for the wrong reason teaches the
  // next reader to widen it until it goes green.
  const paren = masked.indexOf("(", start);
  expect(paren, `${signature} has a parameter list`).toBeGreaterThan(-1);
  let parens = 0;
  let afterParams = -1;
  for (let i = paren; i < masked.length; i++) {
    if (masked[i] === "(") parens++;
    else if (masked[i] === ")" && --parens === 0) {
      afterParams = i;
      break;
    }
  }
  expect(afterParams, `${signature} closes its parameter list`).toBeGreaterThan(-1);
  const open = masked.indexOf("{", afterParams);
  expect(open, `${signature} has a body`).toBeGreaterThan(-1);
  let depth = 0;
  for (let i = open; i < masked.length; i++) {
    if (masked[i] === "{") depth++;
    else if (masked[i] === "}" && --depth === 0) return src.slice(start, i + 1);
  }
  throw new Error(`${signature} never closes its body`);
}

/**
 * The file with its prose removed, so a guard reads what the code DOES.
 *
 * The first version of this matched the words "Office.js" and "PowerPoint.run"
 * in the comments that explain WHY the engine avoids them, so it failed on four
 * files that were entirely correct. A guard that goes red for the wrong reason
 * teaches the next reader to widen it until it goes green.
 *
 * The stripper is shared now — `scripts/without-prose.mjs` — because the same
 * mistake has been made three times in three syntaxes, and three private copies
 * is three chances to write a fourth.
 */
function codeOf(file: string): string {
  return withoutTsProse(readFileSync(file, "utf8")) as string;
}

describe("src/core", () => {
  it("imports nothing from Office.js", () => {
    // The seam the whole design rests on. The engine takes bytes and records
    // and returns bytes, which is what lets it run in the pane, in a CLI and in
    // this suite with no PowerPoint anywhere. One import here and the merge can
    // only ever be tested inside a host that is documented to lie about ids.
    const offenders = filesUnder("src/core").filter((f) => codeOf(f).match(/\bOffice\.|\bPowerPoint\.|office-js/));
    expect(offenders).toEqual([]);
  });

  it("keeps Office.js out of src/host too, so every decision stays testable", () => {
    // src/host is the DECISIONS: which version floor the host clears, where the
    // template's bytes come from, whether an insert did what it said, which
    // slides an undo may remove. src/office is the calls. The whole reason the
    // probe could answer anything is that its judgements lived outside the
    // PowerPoint.run callback and could be checked in CI; one Office.js import
    // here and the same rules become untestable again.
    const offenders = filesUnder("src/host").filter((f) => codeOf(f).match(/\bOffice\.|\bPowerPoint\.|office-js/));
    expect(offenders).toEqual([]);
  });

  it("keeps the decisions out of src/office, so none of them hide in a callback", () => {
    // The other direction, and the one that rots quietly: a rule reimplemented
    // inline next to the call it guards looks tidier and is untestable. Every
    // judgement src/office needs is imported — the version floor, the source
    // choice, the insert reading and the sweep plan from src/host, the plan and
    // the block preparation from src/core — so a file here importing neither is
    // either trivial or has started deciding for itself.
    //
    // This said `src/host` alone until the merge run arrived, and refused it
    // for taking its decisions from `src/core` instead. That was the guard
    // being narrower than its own reason, not the file being wrong.
    const office = filesUnder("src/office");
    expect(office.length).toBeGreaterThan(0);
    for (const file of office) {
      // The RAW source, not codeOf: an import specifier IS a string literal,
      // and codeOf strips those so the Office.js check above cannot trip over
      // prose. Two questions, two readings of the same file.
      expect(readFileSync(file, "utf8"), `${file} decides nothing on its own`).toMatch(/from "\.\.\/(host|core)\//);
    }
  });

  it("lets only the pane's entry point touch Office.js", () => {
    // steps.ts, summary.ts and render.ts are the pane's decisions and its
    // screen, and all three are checked in the suite without a PowerPoint. The
    // moment one of them reads Office.context the labels a user acts on become
    // untestable — and the pane is the surface where a wrong label is the thing
    // that gets pressed.
    const offenders = filesUnder("src/pane")
      .filter((f) => !f.endsWith("main.ts"))
      .filter((f) => codeOf(f).match(/\bOffice\.|\bPowerPoint\./));
    expect(offenders).toEqual([]);
  });

  it("never lets a host call that CHANGES the deck escape without a re-count", () => {
    // src/office cannot run in the suite, so this is a source scan — the same
    // shape PowerChart uses for lockstep it cannot execute.
    //
    // On this host a call can raise and still have done the work: the probe's
    // third sheet timed out on an insert whose deck delta showed both slides
    // had landed. `undoInsert` let that rejection escape, so the confirming
    // re-count never ran and the caller was told the undo had failed while the
    // user's slides were already gone, with no count of what went. `insertDeck`
    // thirty lines above had always caught its own.
    //
    // The rule: a `withTimeout` around a mutating batch sits inside a `try`,
    // and the deck is counted again afterwards.
    const src = readFileSync("src/office/powerpoint.ts", "utf8");
    for (const fn of ["insertDeck", "undoInsert"]) {
      const start = src.indexOf(`export async function ${fn}`);
      expect(start, fn).toBeGreaterThan(-1);
      const next = src.indexOf("\nexport ", start + 1);
      const body = src.slice(start, next === -1 ? undefined : next);
      expect(body, `${fn} catches its own raise`).toMatch(/}\s*catch\s*\(/);
      expect(body, `${fn} counts the deck again afterwards`).toContain("await slideCount()");
    }
  });

  it("never builds a slide id, it asks the host for one", () => {
    // A slide id on this host looks like "256#3561048925". The first merge run
    // counted instead — `for (let n = from; n <= to; n++) ids.push(String(n))`
    // — and handed `["4", "5", "6"]` to `exportAsBase64Presentation`, whose
    // typings say it throws InvalidArgument for an id not in the collection.
    // Both sides were `string`, so tsc saw nothing and all three answer sheets
    // report PowerPointApi 1.10, which means `chooseDeckSource` returns
    // `subset` on the owner's own host and the first press of the merge button
    // would have thrown.
    //
    // The signature is the main guard: `readTemplate` takes slide NUMBERS, so
    // no caller can pass ids and tsc refuses the old call outright. This is the
    // other half — the ids it uses have to come off a loaded collection and
    // through `blockIds`, which is where the checking lives.
    const src = readFileSync("src/office/powerpoint.ts", "utf8");
    const body = functionBody(src, "export async function readTemplate");
    expect(body, "asks the host for the ids").toContain('load("items/id")');
    expect(body, "checks them in src/host rather than here").toContain("blockIds(");
    // `exportAsBase64Presentation` is handed what `blockIds` returned and
    // nothing else. Anything built locally is the defect coming back.
    expect(body).toMatch(/exportAsBase64Presentation\(chosen\.ids\)/);
    // Scoped to this function, not the file. `String(e)` is how the two
    // mutating calls coerce a raise they caught, and a check wide enough to
    // catch that is a check that goes red for the wrong reason — which this
    // file already has a paragraph about.
    expect(body, "builds no id of its own").not.toContain("String(");
  });

  it("looks values up without walking the prototype chain", () => {
    // A field called __proto__ or constructor is a legal spreadsheet header and
    // arrives from a file the user pasted.
    const src = readFileSync("src/core/data/recordset.ts", "utf8");
    expect(src).toContain("Object.create(null)");
  });
});
