import { markdownToSlides } from "./slide-model.ts";
import { createZip, type ZipFileEntry } from "./zip-io.ts";
import { collectDiagnostics } from "./diagnostics.ts";
import { createMediaManager, type MediaManager } from "./media.ts";
import {
  createHyperlinkRel,
  createImageRel,
  DRAWING_NS,
  PRESENTATION_NS,
  REL_NS,
  relsXml,
  type SlideRelationship,
  xmlEscape
} from "./ooxml.ts";
import {
  notesMasterXml,
  notesSlideRelsXml,
  presPropsXml,
  slideLayoutXml,
  slideMasterXml,
  tableStylesXml,
  themeXml,
  viewPropsXml
} from "./pptx-static-parts.ts";
import type {
  ImageSlideBlock,
  MarkdownToPptxOptions,
  MarkdownToPptxResult,
  SlideModel,
  TableCell,
  TextRun
} from "./types.ts";

export { type MarkdownToPptxDiagnostic, type MarkdownToPptxOptions, type MarkdownToPptxResult, type SlideModel } from "./types.ts";
export { markdownToSlides } from "./slide-model.ts";

interface SlideXmlResult {
  xml: string;
  relationships: SlideRelationship[];
}

function parseListRuns(runs: TextRun[]): { runs: TextRun[]; bullet: boolean; level: number } {
  const text = runs.map((run) => run.text).join("");
  const match = text.match(/^(\s*)-\s+(.*)$/);
  if (!match) {
    return { runs, bullet: false, level: 0 };
  }
  const prefixLength = match[1].length + 2;
  let remainingPrefix = prefixLength;
  const trimmedRuns: TextRun[] = [];
  for (const run of runs) {
    if (remainingPrefix >= run.text.length) {
      remainingPrefix -= run.text.length;
      continue;
    }
    const textPart = remainingPrefix > 0 ? run.text.slice(remainingPrefix) : run.text;
    remainingPrefix = 0;
    trimmedRuns.push(run.href ? { text: textPart, href: run.href } : { text: textPart });
  }
  return {
    runs: trimmedRuns,
    bullet: true,
    level: Math.min(8, Math.floor(match[1].length / 2))
  };
}

function textRunXml(run: TextRun, index: number, relationships: SlideRelationship[]): string {
  const relId = run.href ? createHyperlinkRel(relationships, run.href) : undefined;
  const hyperlink = relId ? `<a:hlinkClick r:id="${relId}"/>` : "";
  return `<a:r><a:rPr lang="en-US" sz="${index === 0 ? 2400 : 1800}">${hyperlink}</a:rPr><a:t>${xmlEscape(run.text)}</a:t></a:r>`;
}

function textParagraphFromRuns(runs: TextRun[], index: number, relationships: SlideRelationship[]): string {
  const parsed = parseListRuns(runs);
  const pPr = parsed.bullet ? `<a:pPr${parsed.level > 0 ? ` lvl="${parsed.level}"` : ""}><a:buChar char="•"/></a:pPr>` : "";
  const body = parsed.runs.length > 0 ? parsed.runs.map((run) => textRunXml(run, index, relationships)).join("") : textRunXml({ text: " " }, index, relationships);
  return `<a:p>${pPr}${body}<a:endParaRPr lang="en-US"/></a:p>`;
}

function textParagraph(text: string, index: number, relationships: SlideRelationship[]): string {
  return textParagraphFromRuns([{ text }], index, relationships);
}

function tableXml(rows: TableCell[][], id: number, relationships: SlideRelationship[]): string {
  const columnCount = Math.max(1, ...rows.map((row) => row.length));
  const normalizedRows = rows.length > 0 ? rows : [[{ text: "", runs: [] }]];
  const gridColumns = Array.from({ length: columnCount }, () => '<a:gridCol w="1828800"/>').join("");
  const tableRowsXml = normalizedRows.map((row) => {
    const cells = Array.from({ length: columnCount }, (_, index) => row[index] ?? { text: "", runs: [] });
    return `<a:tr h="370840">${cells.map((cell) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${textParagraphFromRuns(cell.runs.length > 0 ? cell.runs : [{ text: cell.text }], 1, relationships)}</a:txBody><a:tcPr/></a:tc>`).join("")}</a:tr>`;
  }).join("");
  return `<p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="685800" y="2743200"/><a:ext cx="7772400" cy="1828800"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>${gridColumns}</a:tblGrid>${tableRowsXml}</a:tbl></a:graphicData></a:graphic>
      </p:graphicFrame>`;
}

function pictureXml(image: ImageSlideBlock, id: number, relId: string): string {
  const name = image.altText || image.url || `Image ${id}`;
  const descr = image.altText || image.url || "";
  return `<p:pic>
        <p:nvPicPr><p:cNvPr id="${id}" name="${xmlEscape(name)}" descr="${xmlEscape(descr)}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
        <p:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr><a:xfrm><a:off x="685800" y="3886200"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      </p:pic>`;
}

function notesXml(slide: SlideModel): string {
  const relationships: SlideRelationship[] = [];
  const noteParagraphs = slide.notes
    .map((runs, index) => textParagraphFromRuns(runs, index + 1, relationships))
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="${DRAWING_NS}" xmlns:r="${REL_NS}" xmlns:p="${PRESENTATION_NS}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Slide image placeholder"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Notes placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/>${noteParagraphs}</p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:notes>`;
}

function slideXml(slide: SlideModel, index: number, media: MediaManager): SlideXmlResult {
  const relationships: SlideRelationship[] = [{
    id: "rId1",
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
    target: "../slideLayouts/slideLayout1.xml"
  }];
  const textBlocks = slide.blocks.filter((block) => block.kind === "text");
  const tableBlocks = slide.blocks.filter((block) => block.kind === "table");
  const imageBlocks = slide.blocks.filter((block) => block.kind === "image");
  const bodyBlocks = textBlocks.length > 0 ? textBlocks : [{ kind: "text" as const, text: " ", runs: [{ text: " " }] }];
  if (slide.notes.length > 0) {
    relationships.push({
      id: `rId${relationships.length + 1}`,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
      target: `../notesSlides/notesSlide${index}.xml`
    });
  }
  const imageXml = imageBlocks.map((image, imageIndex) => {
    const added = media.addImage(image);
    if (!added) {
      return "";
    }
    const relId = createImageRel(relationships, added.target);
    return pictureXml(image, 40 + imageIndex, relId);
  }).join("\n");
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${DRAWING_NS}" xmlns:r="${REL_NS}" xmlns:p="${PRESENTATION_NS}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title ${index}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="685800" y="457200"/><a:ext cx="7772400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/>${textParagraph(slide.title, 0, relationships)}</p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Body ${index}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="685800" y="1600200"/><a:ext cx="7772400" cy="4572000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>${bodyBlocks.map((block, lineIndex) => textParagraphFromRuns(block.runs, lineIndex + 1, relationships)).join("")}</p:txBody>
      </p:sp>
      ${tableBlocks.map((block, tableIndex) => tableXml(block.rows, 4 + tableIndex, relationships)).join("\n")}
      ${imageXml}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
  return { xml, relationships };
}

function presentationXml(slides: SlideModel[]): string {
  const hasNotes = slides.some((slide) => slide.notes.length > 0);
  const notesMasterRelId = `rId${slides.length + 2}`;
  const defaultTextStyle = `<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${DRAWING_NS}" xmlns:r="${REL_NS}" xmlns:p="${PRESENTATION_NS}">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  ${hasNotes ? `<p:notesMasterIdLst><p:notesMasterId r:id="${notesMasterRelId}"/></p:notesMasterIdLst>` : ""}
  <p:sldIdLst>
${slides.map((_, index) => `    <p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`).join("\n")}
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  ${defaultTextStyle}
</p:presentation>`;
}

function contentTypes(slides: SlideModel[]): string {
  const slideOverrides = Array.from({ length: slides.length }, (_, index) =>
    `  <Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
  ).join("\n");
  const notesOverrides = slides
    .map((slide, index) => slide.notes.length > 0
      ? `  <Override PartName="/ppt/notesSlides/notesSlide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`
      : "")
    .filter(Boolean)
    .join("\n");
  const notesMasterOverride = slides.some((slide) => slide.notes.length > 0)
    ? '  <Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>'
    : "";
  const notesSupportOverrides = slides.some((slide) => slide.notes.length > 0)
    ? [
      '  <Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
      '  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>',
      '  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>',
      '  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>'
    ].join("\n")
    : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="gif" ContentType="image/gif"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
${slideOverrides}
${notesOverrides}
${notesMasterOverride}
${notesSupportOverrides}
</Types>`;
}

export function markdownToPptxResult(markdown: string, options: MarkdownToPptxOptions = {}): MarkdownToPptxResult {
  const slides = markdownToSlides(markdown, options);
  const diagnostics = collectDiagnostics(slides, options);
  const media = createMediaManager(options, diagnostics);
  const hasNotes = slides.some((slide) => slide.notes.length > 0);
  const entries: ZipFileEntry[] = [
    { path: "[Content_Types].xml", data: contentTypes(slides) },
    { path: "_rels/.rels", data: relsXml([
      { id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "ppt/presentation.xml" },
      { id: "rId2", type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", target: "docProps/core.xml" },
      { id: "rId3", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", target: "docProps/app.xml" }
    ]) },
    { path: "docProps/app.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>miku-md2pptx</Application><Slides>${slides.length}</Slides></Properties>` },
    { path: "docProps/core.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${xmlEscape(options.title ?? slides[0]?.title ?? "Markdown deck")}</dc:title><dc:creator>miku-md2pptx</dc:creator></cp:coreProperties>` },
    { path: "ppt/presentation.xml", data: presentationXml(slides) },
    { path: "ppt/_rels/presentation.xml.rels", data: relsXml([
      { id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", target: "slideMasters/slideMaster1.xml" },
      ...slides.map((_, index) => ({ id: `rId${index + 2}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", target: `slides/slide${index + 1}.xml` })),
      ...(hasNotes ? [
        { id: `rId${slides.length + 2}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster", target: "notesMasters/notesMaster1.xml" },
        { id: `rId${slides.length + 3}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps", target: "presProps.xml" },
        { id: `rId${slides.length + 4}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps", target: "viewProps.xml" },
        { id: `rId${slides.length + 5}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", target: "theme/theme1.xml" },
        { id: `rId${slides.length + 6}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles", target: "tableStyles.xml" }
      ] : [])
    ]) },
    { path: "ppt/slideMasters/slideMaster1.xml", data: slideMasterXml() },
    { path: "ppt/slideMasters/_rels/slideMaster1.xml.rels", data: relsXml([
      { id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", target: "../slideLayouts/slideLayout1.xml" },
      { id: "rId2", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", target: "../theme/theme1.xml" }
    ]) },
    { path: "ppt/slideLayouts/slideLayout1.xml", data: slideLayoutXml() },
    { path: "ppt/slideLayouts/_rels/slideLayout1.xml.rels", data: relsXml([
      { id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", target: "../slideMasters/slideMaster1.xml" }
    ]) },
    { path: "ppt/theme/theme1.xml", data: themeXml() }
  ];

  for (const [index, slide] of slides.entries()) {
    const slidePart = slideXml(slide, index + 1, media);
    entries.push({ path: `ppt/slides/slide${index + 1}.xml`, data: slidePart.xml });
    entries.push({ path: `ppt/slides/_rels/slide${index + 1}.xml.rels`, data: relsXml(slidePart.relationships) });
    if (slide.notes.length > 0) {
      entries.push({ path: `ppt/notesSlides/notesSlide${index + 1}.xml`, data: notesXml(slide) });
      entries.push({ path: `ppt/notesSlides/_rels/notesSlide${index + 1}.xml.rels`, data: notesSlideRelsXml(index + 1) });
    }
  }

  if (hasNotes) {
    entries.push({ path: "ppt/notesMasters/notesMaster1.xml", data: notesMasterXml() });
    entries.push({ path: "ppt/notesMasters/_rels/notesMaster1.xml.rels", data: relsXml([
      { id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", target: "../theme/theme2.xml" }
    ]) });
    entries.push({ path: "ppt/theme/theme2.xml", data: themeXml() });
    entries.push({ path: "ppt/presProps.xml", data: presPropsXml() });
    entries.push({ path: "ppt/viewProps.xml", data: viewPropsXml() });
    entries.push({ path: "ppt/tableStyles.xml", data: tableStylesXml() });
  }

  for (const entry of media.entries) {
    entries.push(entry);
  }

  return {
    pptx: createZip(entries),
    diagnostics
  };
}

export function markdownToPptx(markdown: string, options: MarkdownToPptxOptions = {}): Uint8Array {
  return markdownToPptxResult(markdown, options).pptx;
}
