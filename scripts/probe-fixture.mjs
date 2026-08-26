/**
 * The one-slide deck the probe's fixtures are built from.
 *
 * Minimal on purpose: the question is whether PowerPoint accepts an inserted
 * slide, not whether it renders a chart, and every byte here has to travel
 * inside a snippet somebody pastes into an editor.
 */
import JSZip from "jszip";

const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
const T = {
  slide: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
  master: "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
  layout: "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
  theme: "application/vnd.openxmlformats-officedocument.theme+xml",
  pres: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
};
const RT = {
  doc: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  slide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
  master: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
  layout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
};

export async function makeDeck() {
  const zip = new JSZip();
  const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

  zip.file(
    "[Content_Types].xml",
    `${head}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/ppt/presentation.xml" ContentType="${T.pres}"/>` +
      `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${T.master}"/>` +
      `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${T.layout}"/>` +
      `<Override PartName="/ppt/theme/theme1.xml" ContentType="${T.theme}"/>` +
      `<Override PartName="/ppt/slides/slide1.xml" ContentType="${T.slide}"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${head}<Relationships ${REL}><Relationship Id="rId1" Type="${RT.doc}" Target="ppt/presentation.xml"/></Relationships>`,
  );
  zip.file(
    "ppt/presentation.xml",
    `${head}<p:presentation ${P} ${A} ${R}>` +
      `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
      `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>` +
      `<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    `${head}<Relationships ${REL}>` +
      `<Relationship Id="rId1" Type="${RT.master}" Target="slideMasters/slideMaster1.xml"/>` +
      `<Relationship Id="rId2" Type="${RT.slide}" Target="slides/slide1.xml"/></Relationships>`,
  );
  zip.file(
    "ppt/slideMasters/slideMaster1.xml",
    `${head}<p:sldMaster ${P} ${A} ${R}><p:cSld><p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
      `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
      `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`,
  );
  zip.file(
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    `${head}<Relationships ${REL}>` +
      `<Relationship Id="rId1" Type="${RT.layout}" Target="../slideLayouts/slideLayout1.xml"/>` +
      `<Relationship Id="rId2" Type="${RT.theme}" Target="../theme/theme1.xml"/></Relationships>`,
  );
  zip.file(
    "ppt/slideLayouts/slideLayout1.xml",
    `${head}<p:sldLayout ${P} ${A} ${R} type="blank"><p:cSld><p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
      `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
  );
  zip.file(
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    `${head}<Relationships ${REL}><Relationship Id="rId1" Type="${RT.master}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
  );
  zip.file("ppt/theme/theme1.xml", `${head}<a:theme ${A} name="Probe"><a:themeElements/></a:theme>`);
  zip.file(
    "ppt/slides/slide1.xml",
    `${head}<p:sld ${P} ${A} ${R}><p:cSld><p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Probe"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>SSF Merge probe slide</a:t></a:r></a:p></p:txBody>` +
      `</p:sp></p:spTree>` +
      `<p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">` +
      `<p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="777777"/>` +
      `</p:ext></p:extLst></p:cSld>` +
      `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
  );
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    `${head}<Relationships ${REL}><Relationship Id="rId1" Type="${RT.layout}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
  );
  return zip.generateAsync({ type: "uint8array" });
}
