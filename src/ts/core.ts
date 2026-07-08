import { markdownToSlides } from "./slide-model.ts";
import { createZip, readZipEntries, type ZipFileEntry } from "./zip-io.ts";
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

export const mikuMd2PptxMetadata = {
  productName: "miku-md2pptx",
  artifactRole: "markdown-to-pptx-runtime",
  primaryInput: "markdown",
  primaryOutput: "pptx",
  coreApi: ["markdownToSlides", "markdownToPptx", "markdownToPptxResult"]
} as const;

interface SlideXmlResult {
  xml: string;
  relationships: SlideRelationship[];
}

interface Rect {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

interface TemplateContext {
  entries: Map<string, Uint8Array>;
  layoutPath: string;
  layoutName: string;
  layoutTarget: string;
  titlePlaceholderXml: string;
  bodyPlaceholderXml: string;
  bodyPlaceholderRect?: Rect;
  reason: string;
  presentationXml: string;
  presentationRelsXml: string;
  nextImageIndex: number;
}

const decoder = new TextDecoder();

function readTextEntry(entries: Map<string, Uint8Array>, path: string): string | undefined {
  const data = entries.get(path);
  return data ? decoder.decode(data) : undefined;
}

function getAttribute(tag: string, localName: string): string | undefined {
  const pattern = new RegExp(`(?:^|\\s)(?:[^\\s:=]+:)?${localName}="([^"]*)"`);
  return tag.match(pattern)?.[1];
}

function normalizePackagePath(baseDir: string, target: string): string {
  const parts = `${baseDir}/${target}`.split("/");
  const normalized: string[] = [];
  for (const part of parts) {
    if (!part || part === ".") {
      continue;
    }
    if (part === "..") {
      normalized.pop();
      continue;
    } else {
      normalized.push(part);
    }
  }
  return normalized.join("/");
}

function collectRelationshipTargets(xml: string): SlideRelationship[] {
  return Array.from(xml.matchAll(/<[^<\s:]*:?Relationship\b[^>]*>/g), (match) => {
    const tag = match[0];
    return {
      id: getAttribute(tag, "Id") || "",
      type: getAttribute(tag, "Type") || "",
      target: getAttribute(tag, "Target") || "",
      ...(getAttribute(tag, "TargetMode") === "External" ? { targetMode: "External" as const } : {})
    };
  }).filter((rel) => rel.id && rel.type && rel.target);
}

function hasPlaceholderType(xml: string, types: string[]): boolean {
  return Array.from(xml.matchAll(/<[^<\s:]*:?ph\b[^>]*>/g), (match) => {
    const type = getAttribute(match[0], "type") || "body";
    return types.includes(type);
  }).some(Boolean);
}

function collectTagBlocks(xml: string, localName: string): string[] {
  const pattern = new RegExp(`<[^<\\s:]*:?${localName}\\b[\\s\\S]*?<\\/[^<\\s:]*:?${localName}>`, "g");
  return Array.from(xml.matchAll(pattern), (match) => match[0]);
}

function getPlaceholderTag(shapeXml: string): string | undefined {
  return shapeXml.match(/<[^<\s:]*:?ph\b[^>]*\/?>/)?.[0];
}

function findPlaceholderTag(xml: string, types: string[]): string | undefined {
  for (const shapeXml of collectTagBlocks(xml, "sp")) {
    const placeholder = getPlaceholderTag(shapeXml);
    if (!placeholder) {
      continue;
    }
    const type = getAttribute(placeholder, "type") || "body";
    if (types.includes(type)) {
      return placeholder.endsWith("/>") ? placeholder : `${placeholder.slice(0, -1)}/>`;
    }
  }
  return undefined;
}

function getPlaceholderType(placeholderXml: string): string {
  return getAttribute(placeholderXml, "type") || "body";
}

function findPlaceholderShape(xml: string, placeholderXml: string): string | undefined {
  const targetIdx = getAttribute(placeholderXml, "idx");
  const targetType = getPlaceholderType(placeholderXml);
  for (const shapeXml of collectTagBlocks(xml, "sp")) {
    const candidate = getPlaceholderTag(shapeXml);
    if (!candidate) {
      continue;
    }
    const candidateIdx = getAttribute(candidate, "idx");
    const candidateType = getPlaceholderType(candidate);
    if (targetIdx && candidateIdx === targetIdx) {
      return shapeXml;
    }
    if (!targetIdx && candidateType === targetType) {
      return shapeXml;
    }
    if (targetType === "body" && candidateType === "obj") {
      return shapeXml;
    }
  }
  return undefined;
}

function extractShapeRect(shapeXml: string): Rect | undefined {
  const xfrm = shapeXml.match(/<[^<\s:]*:?xfrm\b[\s\S]*?<\/[^<\s:]*:?xfrm>/)?.[0];
  if (!xfrm) {
    return undefined;
  }
  const offTag = xfrm.match(/<[^<\s:]*:?off\b[^>]*>/)?.[0];
  const extTag = xfrm.match(/<[^<\s:]*:?ext\b[^>]*>/)?.[0];
  const x = offTag ? Number(getAttribute(offTag, "x")) : NaN;
  const y = offTag ? Number(getAttribute(offTag, "y")) : NaN;
  const cx = extTag ? Number(getAttribute(extTag, "cx")) : NaN;
  const cy = extTag ? Number(getAttribute(extTag, "cy")) : NaN;
  return Number.isFinite(x) && Number.isFinite(y) && Number.isFinite(cx) && Number.isFinite(cy)
    ? { x, y, cx, cy }
    : undefined;
}

function findPlaceholderRect(entries: Map<string, Uint8Array>, layoutPath: string, layoutXml: string, placeholderXml: string): Rect | undefined {
  const layoutShape = findPlaceholderShape(layoutXml, placeholderXml);
  const layoutRect = layoutShape ? extractShapeRect(layoutShape) : undefined;
  if (layoutRect) {
    return layoutRect;
  }

  const relsPath = layoutPath.replace("ppt/slideLayouts/", "ppt/slideLayouts/_rels/") + ".rels";
  const relsXml = readTextEntry(entries, relsPath);
  const masterRel = relsXml
    ? collectRelationshipTargets(relsXml).find((rel) => rel.type.endsWith("/slideMaster"))
    : undefined;
  if (!masterRel) {
    return undefined;
  }
  const masterPath = normalizePackagePath("ppt/slideLayouts", masterRel.target);
  const masterXml = readTextEntry(entries, masterPath);
  if (!masterXml) {
    return undefined;
  }
  const masterShape = findPlaceholderShape(masterXml, placeholderXml);
  return masterShape ? extractShapeRect(masterShape) : undefined;
}

function getLayoutName(xml: string, layoutPath: string): string {
  const tag = xml.match(/<[^<\s:]*:?sldLayout\b[^>]*>/)?.[0];
  return tag ? getAttribute(tag, "name") || layoutPath.split("/").pop() || layoutPath : layoutPath;
}

function findTemplateLayout(entries: Map<string, Uint8Array>): { path: string; name: string; reason: string; titlePlaceholderXml: string; bodyPlaceholderXml: string } | undefined {
  const layouts = Array.from(entries.keys())
    .filter((path) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(path));
  let titleOnly: { path: string; name: string; reason: string; titlePlaceholderXml: string; bodyPlaceholderXml: string } | undefined;

  for (const path of layouts) {
    const xml = readTextEntry(entries, path);
    if (!xml) {
      continue;
    }
    const titlePlaceholderXml = findPlaceholderTag(xml, ["title", "ctrTitle"]);
    const bodyPlaceholderXml = findPlaceholderTag(xml, ["body", "obj"]);
    const layout = {
      path,
      name: getLayoutName(xml, path),
      titlePlaceholderXml: titlePlaceholderXml || '<p:ph type="title"/>',
      bodyPlaceholderXml: bodyPlaceholderXml || '<p:ph type="body"/>',
      reason: bodyPlaceholderXml ? "found title and body placeholders" : "found title placeholder only"
    };
    if (titlePlaceholderXml && bodyPlaceholderXml) {
      return layout;
    }
    if (titlePlaceholderXml && !titleOnly) {
      titleOnly = layout;
    }
  }

  return titleOnly;
}

function nextTemplateImageIndex(entries: Map<string, Uint8Array>): number {
  let maxIndex = 0;
  for (const path of entries.keys()) {
    const match = path.match(/^ppt\/media\/image(\d+)\.(?:png|jpe?g|gif)$/i);
    if (match) {
      maxIndex = Math.max(maxIndex, Number(match[1]));
    }
  }
  return maxIndex + 1;
}

function createTemplateContext(templatePptx: Uint8Array): TemplateContext | undefined {
  const entries = readZipEntries(templatePptx);
  const presentationXml = readTextEntry(entries, "ppt/presentation.xml");
  const presentationRelsXml = readTextEntry(entries, "ppt/_rels/presentation.xml.rels");
  if (!presentationXml || !presentationRelsXml) {
    throw new Error("Template PPTX does not contain required presentation parts.");
  }
  const selected = findTemplateLayout(entries);
  if (!selected) {
    return undefined;
  }
  return {
    entries,
    layoutPath: selected.path,
    layoutName: selected.name,
    layoutTarget: `../slideLayouts/${selected.path.split("/").pop()}`,
    titlePlaceholderXml: selected.titlePlaceholderXml,
    bodyPlaceholderXml: selected.bodyPlaceholderXml,
    bodyPlaceholderRect: findPlaceholderRect(entries, selected.path, readTextEntry(entries, selected.path) || "", selected.bodyPlaceholderXml),
    reason: selected.reason,
    presentationXml,
    presentationRelsXml,
    nextImageIndex: nextTemplateImageIndex(entries)
  };
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

function textRunXml(run: TextRun, index: number, relationships: SlideRelationship[], explicitFontSize: boolean): string {
  const relId = run.href ? createHyperlinkRel(relationships, run.href) : undefined;
  const hyperlink = relId ? `<a:hlinkClick r:id="${relId}"/>` : "";
  const fontSize = explicitFontSize ? ` sz="${index === 0 ? 2400 : 1800}"` : "";
  return `<a:r><a:rPr lang="en-US"${fontSize}>${hyperlink}</a:rPr><a:t>${xmlEscape(run.text)}</a:t></a:r>`;
}

function textParagraphFromRuns(runs: TextRun[], index: number, relationships: SlideRelationship[], explicitFontSize = true): string {
  const parsed = parseListRuns(runs);
  const pPr = parsed.bullet ? `<a:pPr${parsed.level > 0 ? ` lvl="${parsed.level}"` : ""}><a:buChar char="•"/></a:pPr>` : "";
  const body = parsed.runs.length > 0 ? parsed.runs.map((run) => textRunXml(run, index, relationships, explicitFontSize)).join("") : textRunXml({ text: " " }, index, relationships, explicitFontSize);
  return `<a:p>${pPr}${body}<a:endParaRPr lang="en-US"/></a:p>`;
}

function textParagraph(text: string, index: number, relationships: SlideRelationship[], explicitFontSize = true): string {
  return textParagraphFromRuns([{ text }], index, relationships, explicitFontSize);
}

function tableXml(rows: TableCell[][], id: number, relationships: SlideRelationship[], explicitFontSize: boolean, bodyRect?: Rect, precedingTextBlockCount = 0): string {
  const columnCount = Math.max(1, ...rows.map((row) => row.length));
  const normalizedRows = rows.length > 0 ? rows : [[{ text: "", runs: [] }]];
  const gridColumns = Array.from({ length: columnCount }, () => '<a:gridCol w="1828800"/>').join("");
  const tableHeight = Math.max(740000, normalizedRows.length * 370840);
  const tableWidth = bodyRect ? Math.min(bodyRect.cx, Math.max(3600000, Math.floor(bodyRect.cx * 0.55))) : 7772400;
  const tableX = bodyRect?.x ?? 685800;
  const textOffset = precedingTextBlockCount > 0 ? Math.min(bodyRect ? Math.max(0, bodyRect.cy - tableHeight) : 2000000, 300000 + precedingTextBlockCount * 520000) : 0;
  const tableY = (bodyRect?.y ?? 2743200) + textOffset;
  const tableRowsXml = normalizedRows.map((row) => {
    const cells = Array.from({ length: columnCount }, (_, index) => row[index] ?? { text: "", runs: [] });
    return `<a:tr h="370840">${cells.map((cell) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${textParagraphFromRuns(cell.runs.length > 0 ? cell.runs : [{ text: cell.text }], 1, relationships, explicitFontSize)}</a:txBody><a:tcPr/></a:tc>`).join("")}</a:tr>`;
  }).join("");
  return `<p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="${tableX}" y="${tableY}"/><a:ext cx="${tableWidth}" cy="${tableHeight}"/></p:xfrm>
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

function slideXml(
  slide: SlideModel,
  index: number,
  media: MediaManager,
  layoutTarget = "../slideLayouts/slideLayout1.xml",
  explicitFontSize = true,
  placeholders?: { title: string; body: string; bodyRect?: Rect }
): SlideXmlResult {
  const relationships: SlideRelationship[] = [{
    id: "rId1",
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
    target: layoutTarget
  }];
  const textBlocks = slide.blocks.filter((block) => block.kind === "text");
  const tableBlocks = slide.blocks.filter((block) => block.kind === "table");
  const imageBlocks = slide.blocks.filter((block) => block.kind === "image");
  const bodyBlocks = textBlocks.length > 0 ? textBlocks : [{ kind: "text" as const, text: " ", runs: [{ text: " " }] }];
  const titlePlaceholder = placeholders?.title ?? '<p:ph type="title"/>';
  const bodyPlaceholder = placeholders?.body ?? '<p:ph type="body"/>';
  const titleShapePr = placeholders ? "<p:spPr/>" : '<p:spPr><a:xfrm><a:off x="685800" y="457200"/><a:ext cx="7772400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>';
  const bodyShapePr = placeholders ? "<p:spPr/>" : '<p:spPr><a:xfrm><a:off x="685800" y="1600200"/><a:ext cx="7772400" cy="4572000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>';
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
        <p:nvSpPr><p:cNvPr id="2" name="Title ${index}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${titlePlaceholder}</p:nvPr></p:nvSpPr>
        ${titleShapePr}
        <p:txBody><a:bodyPr/><a:lstStyle/>${textParagraph(slide.title, 0, relationships, explicitFontSize)}</p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Body ${index}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${bodyPlaceholder}</p:nvPr></p:nvSpPr>
        ${bodyShapePr}
        <p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>${bodyBlocks.map((block, lineIndex) => textParagraphFromRuns(block.runs, lineIndex + 1, relationships, explicitFontSize)).join("")}</p:txBody>
      </p:sp>
      ${tableBlocks.map((block, tableIndex) => tableXml(block.rows, 4 + tableIndex, relationships, explicitFontSize, placeholders?.bodyRect, textBlocks.length + tableIndex)).join("\n")}
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

function extractXmlBlock(xml: string, localName: string): string | undefined {
  return xml.match(new RegExp(`<[^<\\s:]*:?${localName}\\b[\\s\\S]*?<\\/[^<\\s:]*:?${localName}>`))?.[0];
}

function extractXmlSelfClosing(xml: string, localName: string): string | undefined {
  return xml.match(new RegExp(`<[^<\\s:]*:?${localName}\\b[^>]*/>`))?.[0];
}

function templatePresentationXml(slides: SlideModel[], template: TemplateContext): string {
  const hasNotes = slides.some((slide) => slide.notes.length > 0);
  const masterIds = extractXmlBlock(template.presentationXml, "sldMasterIdLst")
    || '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>';
  const slideSize = extractXmlSelfClosing(template.presentationXml, "sldSz")
    || '<p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>';
  const notesSize = extractXmlSelfClosing(template.presentationXml, "notesSz")
    || '<p:notesSz cx="6858000" cy="9144000"/>';
  const defaultTextStyle = extractXmlBlock(template.presentationXml, "defaultTextStyle") || "";
  const notesMasterIdList = hasNotes
    ? extractXmlBlock(template.presentationXml, "notesMasterIdLst") || '<p:notesMasterIdLst><p:notesMasterId r:id="rIdGeneratedNotesMaster"/></p:notesMasterIdLst>'
    : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${DRAWING_NS}" xmlns:r="${REL_NS}" xmlns:p="${PRESENTATION_NS}">
  ${masterIds}
  ${notesMasterIdList}
  <p:sldIdLst>
${slides.map((_, index) => `    <p:sldId id="${256 + index}" r:id="rIdGeneratedSlide${index + 1}"/>`).join("\n")}
  </p:sldIdLst>
  ${slideSize}
  ${notesSize}
  ${defaultTextStyle}
</p:presentation>`;
}

function templatePresentationRelsXml(slides: SlideModel[], template: TemplateContext): string {
  const baseRels = collectRelationshipTargets(template.presentationRelsXml)
    .filter((rel) => !rel.type.endsWith("/slide") && !rel.type.endsWith("/notesSlide"));
  const hasNotes = slides.some((slide) => slide.notes.length > 0);
  if (hasNotes && !baseRels.some((rel) => rel.type.endsWith("/notesMaster"))) {
    baseRels.push({
      id: "rIdGeneratedNotesMaster",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
      target: "notesMasters/notesMaster1.xml"
    });
  }
  const generatedRels = slides.map((_, index) => ({
    id: `rIdGeneratedSlide${index + 1}`,
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
    target: `slides/slide${index + 1}.xml`
  }));
  return relsXml([...baseRels, ...generatedRels]);
}

function templateContentTypesXml(slides: SlideModel[], template: TemplateContext): string {
  const existing = readTextEntry(template.entries, "[Content_Types].xml");
  if (!existing) {
    return contentTypes(slides);
  }
  const withoutGenerated = existing
    .replace(/\s*<Override\b[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g, "")
    .replace(/\s*<Override\b[^>]*PartName="\/ppt\/notesSlides\/notesSlide\d+\.xml"[^>]*\/>/g, "")
    .replace(/\s*<Override\b[^>]*PartName="\/docProps\/app\.xml"[^>]*\/>/g, "")
    .replace(/\s*<Override\b[^>]*PartName="\/docProps\/core\.xml"[^>]*\/>/g, "");
  const additions = [
    '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    ...Array.from({ length: slides.length }, (_, index) =>
      `  <Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
    ),
    ...slides
      .map((slide, index) => slide.notes.length > 0
      ? `  <Override PartName="/ppt/notesSlides/notesSlide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`
      : "")
      .filter(Boolean)
  ];
  if (slides.some((slide) => slide.notes.length > 0) && !/PartName="\/ppt\/notesMasters\/notesMaster1\.xml"/.test(withoutGenerated)) {
    additions.push('  <Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>');
  }
  const additionXml = additions.join("\n");
  return withoutGenerated.replace(/<\/Types>\s*$/, `${additionXml}\n</Types>`);
}

function shouldCopyTemplateEntry(path: string): boolean {
  if (path === "[Content_Types].xml" || path === "ppt/presentation.xml" || path === "ppt/_rels/presentation.xml.rels") {
    return false;
  }
  if (path === "docProps/app.xml" || path === "docProps/core.xml") {
    return false;
  }
  if (/^ppt\/slides(?:\/|$)/.test(path) || /^ppt\/notesSlides(?:\/|$)/.test(path)) {
    return false;
  }
  return true;
}

function createTemplateBaseEntries(slides: SlideModel[], template: TemplateContext, title: string): ZipFileEntry[] {
  const copied = Array.from(template.entries, ([path, data]) => ({ path, data }))
    .filter((entry) => shouldCopyTemplateEntry(entry.path));
  return [
    { path: "[Content_Types].xml", data: templateContentTypesXml(slides, template) },
    ...copied,
    { path: "docProps/app.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>miku-md2pptx</Application><Slides>${slides.length}</Slides></Properties>` },
    { path: "docProps/core.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${xmlEscape(title)}</dc:title><dc:creator>miku-md2pptx</dc:creator></cp:coreProperties>` },
    { path: "ppt/presentation.xml", data: templatePresentationXml(slides, template) },
    { path: "ppt/_rels/presentation.xml.rels", data: templatePresentationRelsXml(slides, template) }
  ];
}

export function markdownToPptxResult(markdown: string, options: MarkdownToPptxOptions = {}): MarkdownToPptxResult {
  const slides = markdownToSlides(markdown, options);
  const diagnostics = collectDiagnostics(slides, options);
  const title = options.title ?? slides[0]?.title ?? "Markdown deck";
  const template = options.templatePptx ? createTemplateContext(options.templatePptx) : undefined;
  if (options.templatePptx && template) {
    diagnostics.push({
      severity: "info",
      code: "template-layout-selected",
      message: `Template layout selected: ${template.layoutName} (${template.layoutPath}); ${template.reason}.`
    });
  } else if (options.templatePptx && !template) {
    diagnostics.push({
      severity: "warning",
      code: "template-layout-fallback",
      message: "No template slide layout with a title placeholder was found; used the default generated layout."
    });
  }
  const media = createMediaManager(options, diagnostics, template?.nextImageIndex);
  const hasNotes = slides.some((slide) => slide.notes.length > 0);
  const entries: ZipFileEntry[] = template ? createTemplateBaseEntries(slides, template, title) : [
    { path: "[Content_Types].xml", data: contentTypes(slides) },
    { path: "_rels/.rels", data: relsXml([
      { id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "ppt/presentation.xml" },
      { id: "rId2", type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", target: "docProps/core.xml" },
      { id: "rId3", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", target: "docProps/app.xml" }
    ]) },
    { path: "docProps/app.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>miku-md2pptx</Application><Slides>${slides.length}</Slides></Properties>` },
    { path: "docProps/core.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${xmlEscape(title)}</dc:title><dc:creator>miku-md2pptx</dc:creator></cp:coreProperties>` },
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
    const slidePart = slideXml(
      slide,
      index + 1,
      media,
      template?.layoutTarget,
      !template,
      template ? { title: template.titlePlaceholderXml, body: template.bodyPlaceholderXml, bodyRect: template.bodyPlaceholderRect } : undefined
    );
    entries.push({ path: `ppt/slides/slide${index + 1}.xml`, data: slidePart.xml });
    entries.push({ path: `ppt/slides/_rels/slide${index + 1}.xml.rels`, data: relsXml(slidePart.relationships) });
    if (slide.notes.length > 0) {
      entries.push({ path: `ppt/notesSlides/notesSlide${index + 1}.xml`, data: notesXml(slide) });
      entries.push({ path: `ppt/notesSlides/_rels/notesSlide${index + 1}.xml.rels`, data: notesSlideRelsXml(index + 1) });
    }
  }

  const addEntryIfMissing = (entry: ZipFileEntry): void => {
    if (!entries.some((existing) => existing.path === entry.path)) {
      entries.push(entry);
    }
  };

  if (hasNotes) {
    addEntryIfMissing({ path: "ppt/notesMasters/notesMaster1.xml", data: notesMasterXml() });
    addEntryIfMissing({ path: "ppt/notesMasters/_rels/notesMaster1.xml.rels", data: relsXml([
      { id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", target: "../theme/theme2.xml" }
    ]) });
    addEntryIfMissing({ path: "ppt/theme/theme2.xml", data: themeXml() });
    addEntryIfMissing({ path: "ppt/presProps.xml", data: presPropsXml() });
    addEntryIfMissing({ path: "ppt/viewProps.xml", data: viewPropsXml() });
    addEntryIfMissing({ path: "ppt/tableStyles.xml", data: tableStylesXml() });
  }

  for (const entry of media.entries) {
    addEntryIfMissing(entry);
  }

  return {
    pptx: createZip(entries),
    diagnostics
  };
}

export function markdownToPptx(markdown: string, options: MarkdownToPptxOptions = {}): Uint8Array {
  return markdownToPptxResult(markdown, options).pptx;
}
