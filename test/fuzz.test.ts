import { describe, expect, it } from "vitest";
import { applyFormat, numericValue } from "../src/core/data/format.js";
import { detectType, parseDelimited, toRecordSet } from "../src/core/data/recordset.js";
import { fieldsInText, mergeParagraph } from "../src/core/merge/text.js";
import { A_NS, elements, parseXml } from "../src/core/pptx/xml.js";
import { makeResolver } from "../src/core/merge/resolve.js";

/**
 * The engine's text boundaries, over input nobody wrote by hand.
 *
 * Every other test here states a case somebody thought of. This one states the
 * PROPERTIES instead and hands them input drawn from the characters that have
 * actually broken this add-in — an unclosed `{{`, a lone `}`, `__proto__` as a
 * column name, a European "1,5", a byte-order mark, a bare carriage return, an
 * empty format spec — in combinations no fixture enumerates.
 *
 * Seeded, so a failure is reproducible from the number in the message. Fixed
 * seeds rather than the clock, for the reason this repo already keeps a flake
 * hunt: a suite that fails on one run in fifty teaches everyone to press the
 * button again.
 *
 * What it is NOT: a replacement for the cases above. A property test says
 * "nothing threw and the invariant held", which is a weaker claim than "this
 * input produces this sentence". It is here for the class of defect a case
 * cannot reach — the one that needs three unusual things at once.
 */

/**
 * xorshift32. Deterministic, tiny, and good enough to shuffle bytes.
 *
 * SCRAMBLED and warmed up, which is not decoration. Seeded with 1, 2, 3 … the
 * raw generator's first outputs are all near zero — the state is the seed, and
 * one round of shifts on a small integer stays small — so every length drawn
 * from the first call came out at the bottom of its range. The corpus averaged
 * SIX CHARACTERS a string and 306 of 400 of them held no placeholder at all,
 * while the suite reported four passing property tests.
 *
 * Found by asking what the corpus matched rather than whether it passed, which
 * is the only thing that ever finds this.
 */
function rng(seed: number): () => number {
  let x = (Math.imul(seed | 0 || 1, 0x9e37_79b9) ^ 0x5f35_6495) | 0 || 1;
  const step = (): number => {
    x ^= x << 13;
    x ^= x >>> 17;
    x ^= x << 5;
    return (x >>> 0) / 0x1_0000_0000;
  };
  for (let i = 0; i < 8; i++) step();
  return step;
}

/** The characters that have broken something here, plus ordinary ones. */
const ALPHABET = [
  "a",
  "B",
  "9",
  " ",
  "\t",
  "\n",
  "\r",
  ",",
  ";",
  '"',
  "'",
  "{",
  "}",
  "|",
  ":",
  ".",
  "-",
  "+",
  "%",
  " ",
  "﻿",
  "é",
  "日",
  "__proto__",
  "constructor",
  "{{",
  "}}",
  "{{Name}}",
  "1,5",
  "1.500",
  "1e21",
  "0x10",
  "2026-03-01",
  "03/01/2026",
];

function pick<T>(next: () => number, from: readonly T[]): T {
  return from[Math.floor(next() * from.length)] as T;
}

function noise(next: () => number, maxParts: number): string {
  let out = "";
  const parts = Math.floor(next() * maxParts);
  for (let i = 0; i <= parts; i++) out += pick(next, ALPHABET);
  return out;
}

/** Names a header could plausibly carry, hazards included. */
const NAMES = ["First", "City", " Name ", "__proto__", "constructor", "a|b", "é", "日", "9", "{{x", "", " "];

/**
 * The format specs the pane and the manual actually produce, and the ones
 * somebody types by mistake.
 *
 * Drawn deliberately rather than assembled from the alphabet. The first version
 * of this file built specs out of noise and hit a REAL format kind zero times
 * in four hundred draws — so the property "a spec that is not a format returns
 * the value unchanged" was the only branch it ever reached, and the four
 * formatters were tested by nothing. Asking what the corpus matched is the only
 * way that shows up: everything passed.
 */
const SPECS = [
  "upper",
  "lower",
  "number",
  "number:2",
  "number:0",
  "number:-1",
  "number:1e2",
  "number:200",
  "date",
  "date:dd MMM yyyy",
  "date:yyyy-MM-dd 00:00",
  "DATE:dd",
  " Number : 2 ",
  "wat",
  "",
  "upper:x",
];

/** Values a cell can hold that a number reader has to decide about. */
const RAWS = [
  "12",
  "1,5",
  "1.500",
  "1 234,56",
  "-0.4",
  "1e21",
  "0x10",
  "  7  ",
  "n/a",
  "",
  "1,234",
  "٣",
  "12%",
  "(3)",
  "2026-03-01",
  "1.005",
];

/**
 * Text with real placeholders in it, not only characters that could form one.
 *
 * Same finding as `SPECS`: noise alone put a placeholder in 12 of 400 strings,
 * so the tiling invariant — the one every text merge rests on — was checked
 * almost entirely against text holding no fields at all.
 */
function withFields(next: () => number, maxParts: number): string {
  let out = "";
  const parts = Math.floor(next() * maxParts);
  for (let i = 0; i <= parts; i++) {
    const roll = next();
    if (roll < 0.35) {
      const name = pick(next, NAMES);
      const format = next() < 0.4 ? `|${pick(next, SPECS)}` : "";
      // A quarter of them deliberately malformed: an unclosed `{{`, a stray
      // `}`, a brace inside the name. Those are the ones that have frozen a tab.
      const shape = next();
      out += shape < 0.75 ? `{{${name}${format}}}` : shape < 0.9 ? `{{${name}${format}` : `{{${name}{{${format}}}`;
    } else out += pick(next, ALPHABET);
  }
  return out;
}

describe("the text boundaries, over input nobody wrote by hand", () => {
  it("reads placeholders that tile the text, whatever is around them", () => {
    /**
     * The invariant every text merge rests on: the hits are in order, do not
     * overlap, sit inside the string, and the pieces between them plus the
     * placeholders themselves reconstruct the original exactly. `mergeParagraph`
     * rebuilds a paragraph from precisely these offsets, so a hit one character
     * out writes a field's value over a character of the author's text, and one
     * that overlaps its neighbour writes it twice.
     */
    for (let seed = 1; seed <= 400; seed++) {
      const next = rng(seed);
      const text = withFields(next, 24);
      const hits = fieldsInText(text);
      let at = 0;
      let rebuilt = "";
      for (const hit of hits) {
        expect(hit.index, `seed ${seed}: a hit before the one before it`).toBeGreaterThanOrEqual(at);
        expect(hit.length, `seed ${seed}: a zero-length placeholder would not terminate`).toBeGreaterThan(0);
        expect(hit.index + hit.length, `seed ${seed}: a hit past the end of the text`).toBeLessThanOrEqual(text.length);
        // The name is what a column header is matched against, so it may carry
        // neither the braces nor the format with it.
        expect(hit.name, `seed ${seed}`).not.toContain("{{");
        expect(hit.name, `seed ${seed}`).not.toContain("|");
        rebuilt += text.slice(at, hit.index) + text.slice(hit.index, hit.index + hit.length);
        at = hit.index + hit.length;
      }
      expect(rebuilt + text.slice(at), `seed ${seed}: the hits do not tile the text`).toBe(text);
    }
  });

  it("answers a string for any raw value and any format spec", () => {
    /**
     * `applyFormat`'s own contract: a value that does not match its format is
     * returned unchanged. Not blanked, not an error marker, and never a throw —
     * this runs inside the merge, and an exception here is a run that ends with
     * no slides and nothing on screen.
     */
    for (let seed = 1; seed <= 400; seed++) {
      const next = rng(seed * 7919);
      // Half real specs, half noise, so both branches of the contract are
      // exercised rather than only the fall-through.
      const raw = next() < 0.5 ? pick(next, RAWS) : noise(next, 6);
      const spec = next() < 0.6 ? pick(next, SPECS) : noise(next, 4);
      const out = applyFormat(raw, spec);
      expect(typeof out, `seed ${seed}: raw=${JSON.stringify(raw)} spec=${JSON.stringify(spec)}`).toBe("string");
      // A spec naming no format this add-in has is not a format, and the value
      // goes through untouched.
      const kind = (spec.split(":")[0] ?? "").trim().toLowerCase();
      if (!["upper", "lower", "number", "date"].includes(kind))
        expect(out, `seed ${seed}: spec=${JSON.stringify(spec)}`).toBe(raw);
    }
  });

  it("builds a record set whose rows only ever answer for its own columns", () => {
    /**
     * Two properties the merge relies on and nothing states. A row may not
     * carry a key the column list does not name — `row[field]` is how every
     * placeholder is resolved, so a stray key is a field resolving to data
     * nobody said belonged to it. And a name reached through the prototype is
     * the `__proto__` hazard this repo guards at every table: `constructor` as
     * a column header must resolve to the cell or to nothing, never to a
     * function.
     */
    for (let seed = 1; seed <= 300; seed++) {
      const next = rng(seed * 104_729);
      const set = toRecordSet(parseDelimited(withFields(next, 40)));
      const names = new Set(set.columns.map((c) => c.name));
      expect(set.columns.length, `seed ${seed}: a duplicate column name`).toBe(names.size);
      for (const row of set.rows) {
        for (const key of Object.keys(row)) expect(names.has(key), `seed ${seed}: row key ${key}`).toBe(true);
        const resolve = makeResolver(row);
        for (const suspect of ["__proto__", "constructor", "toString"]) {
          expect(typeof resolve(suspect, undefined), `seed ${seed}: ${suspect} reached the prototype`).not.toBe(
            "function",
          );
        }
      }
      for (const column of set.columns) {
        expect(["text", "number", "date", "image"], `seed ${seed}`).toContain(column.type);
        expect(detectType(set.rows.map((r) => r[column.name] ?? "")), `seed ${seed}`).toBeTruthy();
      }
    }
  });

  it("fills a placeholder split across runs exactly as it fills one that is not", () => {
    /**
     * The hazard this whole file exists for. PowerPoint splits a paragraph into
     * runs wherever formatting changes — and wherever it feels like it, which
     * is why `{{Fi}}{{rst}}` shows up in real templates — so `editRuns` maps
     * offsets in the JOINED string back onto several nodes. An off-by-one there
     * writes a value over a character of the author's text or drops one, and
     * the only visible symptom is a slide that reads slightly wrong.
     *
     * The oracle is computed independently: take the hits, replace each one
     * with what the resolver answers, and the merged paragraph's joined text
     * must equal that string — whatever run boundaries the text was cut at,
     * including boundaries inside a placeholder's braces.
     */
    const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    const row = { First: "Ada", City: "Cairo", __proto__: "x", "9": "nine", "a|b": "pipe", "": "blank" };
    for (let seed = 1; seed <= 300; seed++) {
      const next = rng(seed * 2_654_435_761);
      const text = withFields(next, 18);
      // Cut into runs at random points, so a placeholder is routinely split.
      const cuts = [0];
      for (let i = 0; i < 4; i++) cuts.push(Math.floor(next() * (text.length + 1)));
      cuts.push(text.length);
      cuts.sort((x, y) => x - y);
      const parts: string[] = [];
      for (let i = 1; i < cuts.length; i++) parts.push(text.slice(cuts[i - 1], cuts[i]));

      const resolve = makeResolver(row);
      let want = "";
      let at = 0;
      for (const hit of fieldsInText(text)) {
        const answer = resolve(hit.name, hit.format);
        want += text.slice(at, hit.index) + (answer ?? text.slice(hit.index, hit.index + hit.length));
        at = hit.index + hit.length;
      }
      want += text.slice(at);

      // XML normalises a carriage return in a text node to a newline before
      // any reader sees it, so the oracle has to do the same — this is the
      // parser's contract, not the merge's behaviour, and holding the merge to
      // the un-normalised string would be asserting a defect that is not there.
      const normalise = (t: string): string => t.replace(/\r\n?/g, "\n");
      const runs = parts.map((t) => `<a:r><a:t>${escapeXml(t)}</a:t></a:r>`).join("");
      const doc = parseXml(`<a:p ${A}>${runs}</a:p>`);
      const p = doc.documentElement as unknown as Element;
      mergeParagraph(p, resolve);
      // Read back through the DOM, not with a regular expression. The first
      // version scraped `<a:t>…</a:t>` out of the serialised XML and an EMPTY
      // run serialises as `<a:t/>`, which that pattern matched as an opening
      // tag — so it swallowed the markup up to the next close and reported a
      // failure that was entirely its own. An approximate metric invents
      // defects as well as missing them.
      const got = elements(doc, A_NS, "t")
        .map((t) => t.textContent ?? "")
        .join("");
      expect(got, `seed ${seed}: text=${JSON.stringify(text)} cut at ${cuts.join(",")}`).toBe(normalise(want));
    }
  });

  it("never reads a number out of something a spreadsheet would not", () => {
    // `numericValue` decides whether a chart cell is written, so a false
    // positive puts a made-up number in somebody's chart. Anything it answers
    // must be a finite number.
    for (let seed = 1; seed <= 400; seed++) {
      const next = rng(seed * 31 + 7);
      const raw = next() < 0.6 ? pick(next, RAWS) : noise(next, 5);
      const n = numericValue(raw);
      if (n !== undefined) expect(Number.isFinite(n), `seed ${seed}: ${JSON.stringify(raw)} read as ${n}`).toBe(true);
    }
  });
});

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
