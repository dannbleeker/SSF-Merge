import { describe, expect, it } from "vitest";
import { canBeField, fieldsIn, mergeDocument, mergeParagraph } from "../src/core/merge/text.js";
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
});
