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

/**
 * What `RAWS` MEANS, so a false positive is visible.
 *
 * `null` for the ones a spreadsheet would not read as a number.
 */
const KNOWN: Record<string, number | null> = {
  "12": 12,
  "1,5": 1.5,
  "1.500": 1500,
  "-0.4": -0.4,
  "1e21": null,
  "0x10": null,
  "  7  ": 7,
  "n/a": null,
  "": null,
  "1,234": 1234,
  "12%": null,
  "(3)": null,
  // A THOUSAND and five, not one-and-a-bit: `1.005` is a dot group, by the same
  // decided rule that reads `1,500` as fifteen hundred. Pinning it is how this
  // test found out; the first version of this table guessed 1.005.
  "1.005": 1005,
};

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
  return generate(next, maxParts).text;
}

/**
 * The same, saying how many placeholders it emitted WELL FORMED and with a name
 * that is certainly a field.
 *
 * Without that count the tiling property below is a tautology: `fieldsInText`
 * returning `[]` for every input satisfies "the hits are ordered, bounded and
 * reconstruct the text" perfectly, and it did — the property passed against a
 * gutted implementation. A generator that cannot say what it put in cannot hold
 * a reader to finding it.
 *
 * Only the unambiguous ones are counted: a name from `NAMES` that carries a
 * letter or digit, no braces and no pipe, closed properly. Anything else is
 * still generated — the malformed shapes are the ones that have frozen a tab —
 * it simply does not go into the floor.
 */
function generate(next: () => number, maxParts: number): { text: string; certain: number } {
  let out = "";
  let certain = 0;
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
      if (shape < 0.75 && /[\p{L}\p{N}]/u.test(name) && !name.includes("{") && !name.includes("|")) certain++;
    } else out += pick(next, ALPHABET);
  }
  return { text: out, certain };
}

describe("the text boundaries, over input nobody wrote by hand", () => {
  it("reads placeholders that tile the text, whatever is around them", () => {
    let total = 0;
    let put = 0;
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
      const { text, certain } = generate(next, 24);
      const hits = fieldsInText(text);
      total += hits.length;
      put += certain;
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
    // The FLOOR, without which everything above is a tautology: `text.slice(at,
    // hit.index) + text.slice(hit.index, hit.index + hit.length)` is one slice,
    // so once ordering and bounds hold the reconstruction cannot fail — and an
    // empty hit list satisfies every assertion in the loop. `fieldsInText`
    // returning `[]` for every input passed this property until this line.
    //
    // Corpus-wide rather than per seed: a well-formed placeholder can still be
    // absorbed by an unclosed `{{` the alphabet emitted just before it, which
    // is a correct reading and not a miss, so a per-seed floor has false
    // positives. Most of what was put in has to come back.
    expect(put, "the corpus put no placeholders in front of the scanner").toBeGreaterThan(300);
    expect(total, "the scanner found almost none of them").toBeGreaterThan(put * 0.9);
  });

  it("answers a string for any raw value and any format spec", () => {
    let formatted = 0;
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
      const where = `seed ${seed}: raw=${JSON.stringify(raw)} spec=${JSON.stringify(spec)}`;
      expect(typeof out, where).toBe("string");
      // A spec naming no format this add-in has is not a format, and the value
      // goes through untouched.
      const kind = (spec.split(":")[0] ?? "").trim().toLowerCase();
      if (!["upper", "lower", "number", "date"].includes(kind)) expect(out, where).toBe(raw);
      // And the two formatters whose answer is a pure function of the input are
      // ASSERTED, not merely reached. `typeof out === "string"` is satisfied by
      // `applyFormat` returning the cell for every spec — and a version that
      // formatted nothing at all passed this property.
      if (kind === "upper") expect(out, where).toBe(raw.toLocaleUpperCase());
      if (kind === "lower") expect(out, where).toBe(raw.toLocaleLowerCase());
      if (out !== raw) formatted++;
    }
    // And the corpus changed a value often enough for that to mean something.
    expect(formatted, "no draw ever changed a value, so nothing above was tested").toBeGreaterThan(10);
    // The two formatters whose answer is not a pure function of the input get
    // an ANCHOR, because nothing in the loop constrains them: `number` and
    // `date` gutted to return the cell satisfy every assertion above.
    expect(applyFormat("1234567.891", "number:2"), "the number formatter").toBe("1 234 567,89");
    expect(applyFormat("2026-03-01", "date:d MMM yyyy"), "the date formatter").toBe("1 Mar 2026");
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
        // The column's stated type must be the type of the cells the set holds
        // for it. Asserting only that `detectType` answers SOMETHING cannot
        // fail — every `ColumnType` is a non-empty string — and a version
        // returning `"image"` for everything passed this property.
        expect(detectType(set.rows.map((r) => r[column.name] ?? "")), `seed ${seed}: ${column.name}`).toBe(column.type);
      }
    }
    // The ANCHOR. Everything above compares the type a set states against the
    // function that stated it, so the two agree by construction however wrong
    // that function is — a `detectType` answering `"image"` for everything
    // satisfies it. One table whose answer is known from outside the code turns
    // the pair into a statement about typing rather than about agreement.
    expect(
      toRecordSet([
        ["Amount", "When", "Photo", "Note"],
        ["12", "2026-03-01", "ada.png", "n/a"],
        ["7", "2026-04-01", "grace.jpg", "12"],
      ]).columns.map((c) => c.type),
      "a column's type is not read from its cells",
    ).toEqual(["number", "date", "image", "text"]);
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
    // `__proto__` written in an OBJECT LITERAL sets the prototype instead of a
    // key, and a string value makes it a silent no-op — so the hazard this row
    // was meant to carry was not in it at all. Defined explicitly, it is a real
    // own property, which is what a parsed row actually holds.
    const row: Record<string, string> = { First: "Ada", City: "Cairo", "9": "nine", "a|b": "pipe", "": "blank" };
    Object.defineProperty(row, "__proto__", { value: "x", enumerable: true, writable: true, configurable: true });
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
    // The ANCHOR, for the reason the type property has one. The oracle above is
    // built from the same `fieldsInText` and the same `makeResolver` as the
    // code under test, so any CONSISTENT misreading is invisible to it: a
    // scanner that finds every placeholder and reports every name as `""` fills
    // every slide with the wrong column and satisfies every assertion in the
    // loop, and so does a resolver that fills nothing at all.
    const split = parseXml(`<a:p ${A}><a:r><a:t>Dear {{Fi</a:t></a:r><a:r><a:t>rst}} of {{City}}</a:t></a:r></a:p>`);
    mergeParagraph(split.documentElement as unknown as Element, makeResolver(row));
    expect(
      elements(split, A_NS, "t")
        .map((t) => t.textContent ?? "")
        .join(""),
      "a placeholder split across runs, resolved to the column it names",
    ).toBe("Dear Ada of Cairo");
  });

  it("never reads a number out of something a spreadsheet would not", () => {
    let pinned = 0;
    // `numericValue` decides whether a chart cell is written, so a false
    // positive puts a made-up number in somebody's chart. Anything it answers
    // must be a finite number.
    for (let seed = 1; seed <= 400; seed++) {
      const next = rng(seed * 31 + 7);
      const raw = next() < 0.6 ? pick(next, RAWS) : noise(next, 5);
      const n = numericValue(raw);
      if (n !== undefined) expect(Number.isFinite(n), `seed ${seed}: ${JSON.stringify(raw)} read as ${n}`).toBe(true);
      // Finiteness alone is satisfied by a reader that answers 1 for
      // everything, and one did. `KNOWN` pins what the corpus's own values mean,
      // which is what makes a false POSITIVE visible — and a false positive is
      // the failure this property is named for: a made-up number written into
      // somebody's chart.
      if (Object.prototype.hasOwnProperty.call(KNOWN, raw)) {
        expect(n, `seed ${seed}: ${JSON.stringify(raw)}`).toBe(KNOWN[raw] ?? undefined);
        pinned++;
      }
    }
    expect(pinned, "no draw was a value this test knows the meaning of").toBeGreaterThan(50);
  });
});

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
