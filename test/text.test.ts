import { describe, expect, it } from "vitest";
import { fieldsIn, mergeDocument, mergeParagraph } from "../src/core/merge/text.js";
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
