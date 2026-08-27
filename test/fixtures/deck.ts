/**
 * A minimal .pptx built in memory.
 *
 * Hermetic on purpose: every test starts from bytes this file produced, so a
 * failure is about the engine and never about which PowerPoint wrote the
 * fixture. It is small enough to read and structurally real where the engine
 * touches it, which is the parts, the relationships, the content types, the
 * slide id list and the creation id.
 *
 * It is NOT a validity oracle. Nothing here proves PowerPoint would open the
 * file; that is what a round in the real host is for.
 */
import JSZip from "jszip";

const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';

export interface SlideSpec {
  /** Paragraphs, each given as the runs it is split into. Split a placeholder to reproduce the real thing. */
  paragraphs: string[][];
  /** true for stock notes text, or the text itself — a placeholder in it is the point. */
  notes?: boolean | string;
  /**
   * Text for a chart related from this slide, in DrawingML as a real one holds
   * it. A placeholder here is one the engine does NOT merge, which is the case
   * `prepareBlock` has to report rather than pass over.
   */
  chart?: string;
  creationId?: number;
}

function escapeText(t: string): string {
  return t.replace(/&/g, "&amp;").replace(/</g, "&lt;");
}

function runs(parts: string[]): string {
  return parts
    .map(
      (t, i) =>
        `<a:r><a:rPr lang="en-US" b="${i === 0 ? 1 : 0}" dirty="0"/><a:t>${t
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")}</a:t></a:r>`,
    )
    .join("");
}

function slideXml(spec: SlideSpec): string {
  const paras = spec.paragraphs.map((p) => `<a:p>${runs(p)}</a:p>`).join("");
  const creation =
    spec.creationId === undefined
      ? ""
      : `<p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">` +
        `<p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="${spec.creationId}"/>` +
        `</p:ext></p:extLst>`;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<p:sld ${P} ${A} ${R}><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody></p:sp>` +
    `</p:spTree>${creation}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
  );
}

const TYPE = {
  slide: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
  notes: "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
  chart: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
} as const;

const REL_TYPE = {
  slide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
  notes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
  layout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
  master: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  doc: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  chart: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
} as const;

/** Build a deck whose slides are exactly the specs given. */
export async function makeDeck(slides: SlideSpec[]): Promise<Uint8Array> {
  const zip = new JSZip();

  const overrides = slides
    .map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="${TYPE.slide}"/>`)
    .concat(
      slides.flatMap((s, i) => [
        ...(s.notes
          ? [`<Override PartName="/ppt/notesSlides/notesSlide${i + 1}.xml" ContentType="${TYPE.notes}"/>`]
          : []),
        ...(s.chart ? [`<Override PartName="/ppt/charts/chart${i + 1}.xml" ContentType="${TYPE.chart}"/>`] : []),
      ]),
    )
    .join("");

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
      `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
      `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
      `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
      `${overrides}</Types>`,
  );

  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
      `<Relationship Id="rId1" Type="${REL_TYPE.doc}" Target="ppt/presentation.xml"/></Relationships>`,
  );

  const sldIds = slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join("");
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:presentation ${P} ${A} ${R}>` +
      `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
      `<p:sldIdLst>${sldIds}</p:sldIdLst>` +
      `<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
  );

  const presRels = slides
    .map((_, i) => `<Relationship Id="rId${i + 2}" Type="${REL_TYPE.slide}" Target="slides/slide${i + 1}.xml"/>`)
    .join("");
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
      `<Relationship Id="rId1" Type="${REL_TYPE.master}" Target="slideMasters/slideMaster1.xml"/>` +
      `${presRels}</Relationships>`,
  );

  zip.file(
    "ppt/slideMasters/slideMaster1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldMaster ${P} ${A} ${R}>` +
      `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
      `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
      `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`,
  );
  zip.file(
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
      `<Relationship Id="rId1" Type="${REL_TYPE.layout}" Target="../slideLayouts/slideLayout1.xml"/>` +
      `<Relationship Id="rId2" Type="${REL_TYPE.theme}" Target="../theme/theme1.xml"/></Relationships>`,
  );

  zip.file(
    "ppt/slideLayouts/slideLayout1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldLayout ${P} ${A} ${R} type="blank">` +
      `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
      `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
  );
  zip.file(
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
      `<Relationship Id="rId1" Type="${REL_TYPE.master}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
  );

  zip.file(
    "ppt/theme/theme1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<a:theme ${A} name="Test"><a:themeElements/></a:theme>`,
  );

  slides.forEach((spec, i) => {
    const n = i + 1;
    zip.file(`ppt/slides/slide${n}.xml`, slideXml(spec));
    const notesRel = spec.notes
      ? `<Relationship Id="rId2" Type="${REL_TYPE.notes}" Target="../notesSlides/notesSlide${n}.xml"/>`
      : "";
    const chartRel = spec.chart
      ? `<Relationship Id="rId3" Type="${REL_TYPE.chart}" Target="../charts/chart${n}.xml"/>`
      : "";
    zip.file(
      `ppt/slides/_rels/slide${n}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
        `<Relationship Id="rId1" Type="${REL_TYPE.layout}" Target="../slideLayouts/slideLayout1.xml"/>` +
        `${notesRel}${chartRel}</Relationships>`,
    );
    if (spec.chart) {
      // A chart's text is DrawingML, the same `<a:p>`/`<a:t>` the slide uses,
      // which is why `fieldsIn` finds it. Split across two runs on purpose:
      // that is the ordinary state of a placeholder after an edit, and it is
      // what a regex over the raw markup would miss.
      const half = Math.ceil(spec.chart.length / 2);
      zip.file(
        `ppt/charts/chart${n}.xml`,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
          `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ${A}>` +
          `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>` +
          `<a:r><a:rPr lang="en-US"/><a:t>${escapeText(spec.chart.slice(0, half))}</a:t></a:r>` +
          `<a:r><a:rPr lang="en-US"/><a:t>${escapeText(spec.chart.slice(half))}</a:t></a:r>` +
          `</a:p></c:rich></c:tx></c:title></c:chartSpace>`,
      );
    }
    if (spec.notes) {
      zip.file(
        `ppt/notesSlides/notesSlide${n}.xml`,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:notes ${P} ${A} ${R}>` +
          `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${typeof spec.notes === "string" ? escapeText(spec.notes) : `notes for slide ${n}`}</a:t></a:r></a:p></p:txBody>` +
          `</p:sp></p:spTree></p:cSld></p:notes>`,
      );
      zip.file(
        `ppt/notesSlides/_rels/notesSlide${n}.xml.rels`,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ${REL}>` +
          `<Relationship Id="rId1" Type="${REL_TYPE.slide}" Target="../slides/slide${n}.xml"/></Relationships>`,
      );
    }
  });

  return zip.generateAsync({ type: "uint8array" });
}
