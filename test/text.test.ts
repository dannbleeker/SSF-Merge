import { describe, expect, it } from "vitest";
import { applyFormat } from "../src/core/data/format.js";
import { imageMode } from "../src/core/merge/images.js";
import {
  canBeField,
  fieldsIn,
  fieldsInText,
  mergeDocument,
  mergeParagraph,
  whyNotAField,
  type FieldHit,
} from "../src/core/merge/text.js";
import { A_NS, elements, parseXml, serializeXml } from "../src/core/pptx/xml.js";

const A = `xmlns:a="${A_NS}"`;

/** One paragraph, split into the runs given, each run carrying its own properties. */
function paragraph(...parts: string[]): { doc: Document; p: Element } {
  const runs = parts
    .map((t, i) => `<a:r><a:rPr lang="en-US" b="${i}" sz="${1200 + i}"/><a:t>${t}</a:t></a:r>`)
    .join("");
  const doc = parseXml(`<a:p ${A}>${runs}</a:p>`);
  return { doc, p: doc.documentElement };
}

function text(p: Element): string {
  return elements(p, A_NS, "t")
    .map((t) => t.textContent ?? "")
    .join("");
}

describe("mergeParagraph", () => {
  it("replaces a placeholder that PowerPoint split across runs", () => {
    // This is the ordinary case, not the edge case: an edit or a spellcheck
    // pass leaves `{{FirstName}}` stored in two or three pieces.
    const { p } = paragraph("Dear {{Fir", "stName", "}}, welcome");
    expect(mergeParagraph(p, () => "Ada")).toBe(true);
    expect(text(p)).toBe("Dear Ada, welcome");
  });

  it("leaves the run properties of every run exactly as they were", () => {
    const { doc, p } = paragraph("Hello {{Name", "}}!");
    const before = elements(p, A_NS, "rPr").map((r) => serializeXml(r as unknown as Document));
    mergeParagraph(p, () => "Grace");
    const after = elements(doc.documentElement as unknown as Element, A_NS, "rPr").map((r) =>
      serializeXml(r as unknown as Document),
    );
    expect(after).toEqual(before);
  });

  it("puts the value in the first run the placeholder touches", () => {
    // The template author formats the opening brace to choose how the merged
    // value looks. That contract only holds if the value lands in run one.
    const { p } = paragraph("{{Name", "}}");
    mergeParagraph(p, () => "Ada");
    const ts = elements(p, A_NS, "t");
    expect(ts[0]?.textContent).toBe("Ada");
    expect(ts[1]?.textContent).toBe("");
  });

  it("handles two placeholders in one paragraph independently", () => {
    const { p } = paragraph("{{A}} and ", "{{B}}");
    mergeParagraph(p, (n) => (n === "A" ? "first" : "second"));
    expect(text(p)).toBe("first and second");
  });

  it("leaves an unresolved field visible rather than blanking it", () => {
    // A blank slide looks finished. A visible {{Territory}} does not, and that
    // is the point: the author sees their own typo.
    const { p } = paragraph("Region: {{Territory}}");
    expect(mergeParagraph(p, () => null)).toBe(false);
    expect(text(p)).toBe("Region: {{Territory}}");
  });

  it("resolves one field and leaves the unknown one beside it", () => {
    const { p } = paragraph("{{Known}} / {{Unknown}}");
    mergeParagraph(p, (n) => (n === "Known" ? "yes" : null));
    expect(text(p)).toBe("yes / {{Unknown}}");
  });

  it("marks xml:space when the result starts or ends with a space", () => {
    const { p } = paragraph("{{Name}} ");
    mergeParagraph(p, () => "Ada");
    const t = elements(p, A_NS, "t")[0];
    expect(t?.getAttribute("xml:space")).toBe("preserve");
  });

  it("passes the format spec through to the resolver", () => {
    const { p } = paragraph("{{Amount|number:2}}");
    let seen: [string, string | undefined] | undefined;
    mergeParagraph(p, (n, f) => {
      seen = [n, f];
      return "1 234,00";
    });
    expect(seen).toEqual(["Amount", "number:2"]);
  });

  it("reports no change for a paragraph with no placeholder", () => {
    const { p } = paragraph("Just words");
    expect(mergeParagraph(p, () => "x")).toBe(false);
  });
});

describe("mergeDocument", () => {
  it("reaches paragraphs inside a table cell", () => {
    const doc = parseXml(
      `<p:sld xmlns:p="x" ${A}><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>{{Cell}}</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></p:sld>`,
    );
    expect(mergeDocument(doc, () => "value")).toBe(1);
    expect(serializeXml(doc)).toContain("value");
  });
});

describe("fieldsIn", () => {
  it("lists each field once, in the order it first appears", () => {
    const doc = parseXml(`<p:sld xmlns:p="x" ${A}><a:p><a:r><a:t>{{B}} {{A}} {{B}}</a:t></a:r></a:p></p:sld>`);
    expect(fieldsIn(doc)).toEqual(["B", "A"]);
  });

  it("finds a field split across runs, which is the whole point", () => {
    const doc = parseXml(
      `<p:sld xmlns:p="x" ${A}><a:p><a:r><a:t>{{Fir</a:t></a:r><a:r><a:t>stName}}</a:t></a:r></a:p></p:sld>`,
    );
    expect(fieldsIn(doc)).toEqual(["FirstName"]);
  });
});

describe("a placeholder that is not written in English", () => {
  it("merges a field name with Danish letters in it", () => {
    // `[\w.]` is ASCII-only, and stays ASCII-only under the `u` flag, so
    // `{{Beløb}}` and `{{Måned}}` were INVISIBLE: fieldsIn never reported them,
    // the pane could not flag them as unmatched either, and the literal braces
    // were printed on every merged slide. On a product whose first users write
    // Danish that is most of a template.
    const { doc, p } = paragraph("{{Navn}} skylder {{Beløb}} i {{Måned}}");
    expect(fieldsIn(doc)).toEqual(["Navn", "Beløb", "Måned"]);

    const row: Record<string, string> = { Navn: "Ada", Beløb: "1500", Måned: "Marts" };
    mergeParagraph(p, (name) => row[name] ?? null);
    expect(text(p)).toBe("Ada skylder 1500 i Marts");
  });

  it("merges names in other scripts too, and still splits on the format pipe", () => {
    const { doc, p } = paragraph("{{Größe}} {{Πλήθος}} {{Имя|upper}}");
    expect(fieldsIn(doc)).toEqual(["Größe", "Πλήθος", "Имя"]);
    const seen: [string, string | undefined][] = [];
    mergeParagraph(p, (name, format) => {
      seen.push([name, format]);
      return name;
    });
    expect(seen).toEqual([
      ["Größe", undefined],
      ["Πλήθος", undefined],
      ["Имя", "upper"],
    ]);
  });

  it("still refuses a name made of nothing but punctuation or spaces", () => {
    // The widening must not turn every brace pair into a placeholder. A name
    // has to carry at least one letter or digit.
    const { doc } = paragraph("{{ }} {{!!}} {{ - }}");
    expect(fieldsIn(doc)).toEqual([]);
  });
});

describe("a column name with spaces in it", () => {
  /**
   * Reported from a real run, and the defect was the pane and the engine
   * disagreeing about what a field is.
   *
   * `Row Labels`, `Min. of cost` and `Sum of quantity monthly` are the literal
   * default headers of an Excel pivot table — the commonest thing anybody
   * pastes into this add-in. The reader's character class had no space on it,
   * so the pane's own Insert buttons put those tokens on the slide and the
   * read-back reported the slides carried no fields at all. Visibly on the
   * slide, invisible to the engine, and the pane said so as a fact.
   */
  it("reads the headers an Excel pivot table actually produces", () => {
    const { doc } = paragraph("{{Min. of cost}} {{Row Labels}} {{Sum of quantity monthly}}");
    expect(fieldsIn(doc)).toEqual(["Min. of cost", "Row Labels", "Sum of quantity monthly"]);
  });

  it("merges one, keeping the spaces inside the name and not around it", () => {
    const { p } = paragraph("Total: {{ Sum of quantity }} units");
    const seen: string[] = [];
    mergeParagraph(p, (name) => {
      seen.push(name);
      return "42";
    });
    // Trimmed at the edges, so `{{ Name }}` and `{{Name}}` are one field and
    // both match a header the parse has already trimmed.
    expect(seen).toEqual(["Sum of quantity"]);
    expect(text(p)).toBe("Total: 42 units");
  });

  it("still splits on the format pipe, and does not run into the next field", () => {
    const { doc } = paragraph("{{Min. of cost|money}} and {{Row Labels}}");
    expect(fieldsIn(doc)).toEqual(["Min. of cost", "Row Labels"]);
    const { p } = paragraph("{{Min. of cost|money}}");
    const seen: [string, string | undefined][] = [];
    mergeParagraph(p, (name, format) => {
      seen.push([name, format]);
      return "x";
    });
    expect(seen).toEqual([["Min. of cost", "money"]]);
  });
});

describe("a column the pane may not offer as a field", () => {
  /**
   * The guard that stops this defect coming back by another route. The Insert
   * button builds `{{Column}}` and the engine reads it with `FIELD`, and for an
   * hour nothing checked that those two agree.
   *
   * The interesting case is not a header that fails to match — it is one that
   * matches a DIFFERENT, shorter name. `Total|EUR` would put a field called
   * "Total" on the slide, bound to a column that does not exist, silently.
   */
  it("accepts the ordinary ones, spaces and dots included", () => {
    for (const column of ["First", "Row Labels", "Min. of cost", "Beløb", "Sum of quantity monthly", "Q1 2026"]) {
      expect(canBeField(column), column).toBe(true);
    }
  });

  it("refuses one that would read back as a different name, or as none", () => {
    for (const column of ["Total|EUR", "a}}b", "{{x", "", "   ", "!!"]) {
      expect(canBeField(column), column).toBe(false);
    }
  });

  /**
   * WHICH rule, because the pane printed one sentence for all of them.
   *
   * Three of the six refused above have no brace and no pipe in them, and the
   * commonest refusal in real data has neither either: a header cell holding
   * Alt+Enter arrives quoted on the clipboard, parses perfectly, and is a
   * column whose name has a line break in it.
   */
  it("says which rule the name breaks", () => {
    const why = (column: string): string => whyNotAField(column) ?? "";
    expect(whyNotAField("First")).toBeNull();
    expect(why("Total|EUR")).toContain("brace or a pipe");
    expect(why("a}}b")).toContain("brace or a pipe");
    expect(why("Revenue\n(EUR)"), "an Excel header with Alt+Enter in it").toContain("one line");
    expect(why("!!")).toContain("letter or digit");
    expect(why("   ")).toContain("letter or digit");
    expect(why(" First ")).toContain("start or end with a space");
  });

  it("agrees with `canBeField` about every one of them", () => {
    // Defined as each other, so the chip that is missing and the sentence
    // explaining it cannot come apart. Asserted rather than trusted to the
    // one-line definition, which a later reader is free to unpick.
    for (const column of ["First", "Row Labels", "Total|EUR", "a}}b", "{{x", "", "   ", "!!", "a\nb", " x "]) {
      expect(whyNotAField(column) === null, column).toBe(canBeField(column));
    }
  });
});

describe("the placeholder reader is a scan, not a pattern", () => {
  /**
   * The pattern that used to do this is kept HERE, in the test, because the
   * claim being made is a comparison with it: the scanner reads the same
   * placeholders out of the same text. It is the exact source that shipped —
   * two lazy unbounded classes with a required character between them, which is
   * both what made it correct and what made it quadratic.
   */
  const OLD = (): RegExp =>
    new RegExp("\\{\\{\\s*([^{}|\\r\\n]*?[\\p{L}\\p{N}][^{}|\\r\\n]*?)\\s*(?:\\|\\s*([^{}\\r\\n]+?)\\s*)?\\}\\}", "gu");

  const viaPattern = (s: string): FieldHit[] =>
    [...s.matchAll(OLD())].map((h) => ({
      name: h[1] ?? "",
      // The one declared difference, applied to the OLD side so the comparison
      // below can be an equality. A format of nothing but whitespace is now
      // reported as no format; see `formatIsWhitespace` for the proof that
      // neither the text pass nor the image pass can tell the two apart.
      format: h[2] === undefined || h[2].trim() === "" ? undefined : h[2],
      index: h.index,
      length: h[0].length,
    }));

  /**
   * Deliberately weighted at token SHAPES rather than at random characters. A
   * uniform alphabet almost never spells a placeholder — the first version of
   * this corpus put two million strings through and only 2,533 of them produced
   * a hit at all, which is a measurement of the generator and not of the reader.
   */
  const corpus = (): string[] => {
    // `Math.imul`, not `seed * A`: the textbook LCG written in plain JavaScript
    // overflows 2^53 on its first step and loses the low bits, and the sequence
    // then repeats — 40,000 draws produced 5,317 distinct strings where 20,000
    // draws had produced 5,475. A generator that quietly stops generating is the
    // vacuity this block's first assertion exists to catch, and it caught it.
    let seed = 20260830;
    const rnd = (): number => {
      seed = (seed + 0x6d2b79f5) | 0;
      let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
    const pick = (xs: string[]): string => xs[Math.floor(rnd() * xs.length)] ?? "";

    const ws = ["", " ", "  ", "\t", "\n", "\r\n", " \n ", " ", " \t "];
    const names = [
      "a",
      "Name",
      "Row Labels",
      "1",
      "ø",
      "𝔘",
      "a b",
      "a-b",
      "Min. of cost",
      "",
      " ",
      "-",
      "!",
      "a!b",
      "0",
      "..",
      "a\nb",
      "a\tb",
    ];
    const formats = ["", " ", "  ", "\n", " \n", "number:0", "image", "a|b", "date:d MMM", "!", "\t", " x "];
    const junk = ["", "x", "{", "}", "{{", "}}", "|", "\n", " ", "{{a}}", "}}{{", "{{{"];

    const out = new Set<string>();
    // Every atom in every position, so the corners are covered by construction
    // rather than by luck.
    for (const a of [...names, ...ws, ...junk]) {
      for (const b of [...names, ...formats]) {
        out.add(`{{${a}${b}}}`);
        out.add(`{{${a}|${b}}}`);
        out.add(`{{${a}}}${b}{{${a}}}`);
        out.add(`x${a}{{b|${b}}}y`);
        out.add(`{{${a}`);
        out.add(`{{{${a}|${b}}}}`);
      }
    }
    // And random tokens with random junk between them.
    for (let n = 0; n < 40000; n++) {
      let s = "";
      for (let i = 0; i <= Math.floor(rnd() * 3); i++) {
        const body =
          rnd() < 0.5
            ? `${pick(ws)}${pick(names)}${pick(ws)}`
            : `${pick(ws)}${pick(names)}${pick(ws)}|${pick(ws)}${pick(formats)}${pick(ws)}`;
        s += `${pick(junk)}{{${body}}}`;
      }
      out.add(s + pick(junk));
    }
    return [...out];
  };

  it("reads a corpus that actually holds placeholders", () => {
    // The vacuity guard. Every assertion below is satisfied forever by a corpus
    // of strings that are not placeholders, and this suite has shipped exactly
    // that mistake before.
    const withHits = corpus().filter((s) => viaPattern(s).length > 0);
    expect(corpus().length).toBeGreaterThan(25000);
    expect(withHits.length).toBeGreaterThan(15000);
  });

  it("finds the same placeholders the pattern did, in the same places", () => {
    const differ: string[] = [];
    for (const s of corpus()) {
      if (JSON.stringify(fieldsInText(s)) !== JSON.stringify(viaPattern(s))) differ.push(s);
    }
    expect(differ.slice(0, 5)).toEqual([]);
  });

  it("does not walk past a placeholder that opens inside a broken one", () => {
    // `{{a{{b}}` has no field called `a` — the brace ends it — and does have one
    // called `b`. A scan that skipped to the end of the failed span would miss
    // it, which is the one thing a regular expression got right for free.
    expect(fieldsInText("{{a{{b}}").map((f) => f.name)).toEqual(["b"]);
    // Three braces, not four: `{{{Name}}` is the case that separates "try the
    // next `{{`" from "resume after the span we gave up on", because the only
    // `{{` left starts one character in.
    expect(fieldsInText("{{{Name}}").map((f) => f.name)).toEqual(["Name"]);
  });

  it("refuses a lone closing brace inside the braces", () => {
    expect(fieldsInText("{{a}b}}")).toEqual([]);
    expect(fieldsInText("{{a}}}").map((f) => f.length)).toEqual([5]);
  });

  it("reports where each placeholder is, so a merge can edit around it", () => {
    const hits = fieldsInText("Hi {{ First }}, of {{City|upper}}.");
    expect(hits).toEqual([
      { name: "First", format: undefined, index: 3, length: 11 },
      { name: "City", format: "upper", index: 19, length: 14 },
    ]);
  });

  it("still refuses a pipe with nothing after it", () => {
    // Inherited exactly, and the reason it is worth a test rather than a
    // shrug: `{{Name|}}` prints its own braces on every merged slide, and a
    // reader of the scanner would otherwise assume it had been tidied away.
    expect(fieldsInText("{{Name|}}")).toEqual([]);
    expect(fieldsInText("{{Name|\n}}")).toEqual([]);
  });

  it("treats a whitespace-only format as no format", () => {
    // The one declared difference from the pattern, which reported `" "` here.
    expect(fieldsInText("{{Total|  }}")).toEqual([{ name: "Total", format: undefined, index: 0, length: 12 }]);
  });

  it("cannot be seen by either pass that reads a format", () => {
    // What makes the difference above declarable rather than a behaviour
    // change: both consumers of `format` answer a whitespace one exactly as
    // they answer none.
    for (const spec of [" ", "  ", "\t", " \t "]) {
      expect(applyFormat("Ada", spec), spec).toBe(applyFormat("Ada", undefined));
      expect(imageMode(spec), spec).toBe(imageMode(undefined));
    }
  });

  it("reads an unclosed placeholder in linear time", () => {
    // The defect this replaced: `{{` and forty thousand letters took 4.4
    // seconds of synchronous backtracking, inside a deck somebody was sent, on
    // a call the pane's own timeouts cannot interrupt. The budget is two orders
    // of magnitude under the old measurement and two above the new one, so it
    // fails on the pattern and passes on the scan without being a stopwatch.
    for (const bad of [
      "{{" + "a".repeat(40000),
      "{{" + " ".repeat(40000),
      "{{a|" + " ".repeat(4000),
      "{{a".repeat(10000),
    ]) {
      const began = performance.now();
      expect(fieldsInText(bad)).toEqual([]);
      expect(performance.now() - began).toBeLessThan(200);
    }
  });

  it("reads ordinary slide text as fast as the pattern did", () => {
    const ordinary = "{{Name|upper}} lorem ipsum dolor sit amet {{City}} ".repeat(6000);
    const began = performance.now();
    expect(fieldsInText(ordinary)).toHaveLength(12000);
    expect(performance.now() - began).toBeLessThan(500);
  });
});

/**
 * A merge that produces a file PowerPoint will not open at all.
 *
 * XML 1.0 has no representation for most of the C0 controls, for a lone
 * surrogate, or for `FFFE`/`FFFF` — not as a character and not as an entity.
 * `@xmldom/xmldom` writes them straight through and parses them straight back,
 * so every check in this repo passes on a part a conforming parser refuses;
 * PowerPoint is a conforming parser, and it condemns the whole deck.
 *
 * The measurement is made on the SERIALISED part rather than on `textContent`,
 * because textContent is exactly where the round trip hides it. Python's expat
 * is the second opinion: `serializeXml` -> `parseXml` cannot show the defect,
 * since both ends are the lenient parser that caused it.
 */
describe("a cell holding a character XML cannot carry", () => {
  const VERTICAL_TAB = String.fromCharCode(0x0b);
  const NUL = String.fromCharCode(0x00);
  const LONE_SURROGATE = String.fromCharCode(0xd800);

  /** Every code unit in the markup that XML 1.0's `Char` production forbids. */
  function illegal(xml: string): number[] {
    const out: number[] = [];
    for (let i = 0; i < xml.length; i++) {
      const c = xml.charCodeAt(i);
      const control = c < 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d;
      const surrogate = c >= 0xd800 && c <= 0xdfff;
      const paired = surrogate && c < 0xdc00 && xml.charCodeAt(i + 1) >= 0xdc00 && xml.charCodeAt(i + 1) <= 0xdfff;
      const noncharacter = c === 0xfffe || c === 0xffff;
      if ((control || (surrogate && !paired) || noncharacter) && !paired) out.push(c);
      if (paired) i++;
    }
    return out;
  }

  it("cannot put one on a slide, whichever character it is", () => {
    for (const bad of [VERTICAL_TAB, NUL, LONE_SURROGATE, String.fromCharCode(0x1f)]) {
      const { doc, p } = paragraph("Hello {{Na", "me}}");
      mergeParagraph(p, () => "Ada" + bad + "Lovelace");
      expect(illegal(serializeXml(doc))).toEqual([]);
    }
  });

  it("substitutes a space, so the words either side stay apart", () => {
    const { p } = paragraph("Hello {{Name}}");
    mergeParagraph(p, () => "Ada" + VERTICAL_TAB + "Lovelace");
    expect(text(p)).toBe("Hello Ada Lovelace");
  });

  it("leaves an astral character alone — a surrogate PAIR is one legal code point", () => {
    const { doc, p } = paragraph("Hello {{Name}}");
    mergeParagraph(p, () => "Ada \u{1F600}");
    expect(text(p)).toBe("Hello Ada \u{1F600}");
    expect(illegal(serializeXml(doc))).toEqual([]);
  });

  it("leaves tab, newline and carriage return alone — XML carries all three", () => {
    const { p } = paragraph("Hello {{Name}}");
    mergeParagraph(p, () => "a\tb\nc");
    expect(text(p)).toBe("Hello a\tb\nc");
  });
});
