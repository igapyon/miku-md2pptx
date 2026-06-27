export const PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
export const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

export interface SlideRelationship {
  id: string;
  type: string;
  target: string;
  targetMode?: "External";
}

export function xmlEscape(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export function createHyperlinkRel(relationships: SlideRelationship[], href: string): string {
  const existing = relationships.find((rel) => rel.target === href && rel.type.endsWith("/hyperlink"));
  if (existing) {
    return existing.id;
  }
  const id = `rId${relationships.length + 1}`;
  relationships.push({
    id,
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    target: href,
    targetMode: "External"
  });
  return id;
}

export function createImageRel(relationships: SlideRelationship[], target: string): string {
  const id = `rId${relationships.length + 1}`;
  relationships.push({
    id,
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    target
  });
  return id;
}

export function relsXml(relationships: SlideRelationship[]): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${relationships.map((rel) => `  <Relationship Id="${rel.id}" Type="${rel.type}" Target="${xmlEscape(rel.target)}"${rel.targetMode ? ` TargetMode="${rel.targetMode}"` : ""}/>`).join("\n")}
</Relationships>`;
}
