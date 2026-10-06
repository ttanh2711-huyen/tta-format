import type { LineSpacingSetting, ParagraphAlignment } from "../standards/types";
import type { StyleSpec } from "./styleSpecs";

/** Word tự bỏ dấu gạch dưới khi tạo mã style (w:styleId) từ tên style. */
export function styleIdFromName(name: string): string {
  return name.replace(/_/g, "");
}

const JC: Record<ParagraphAlignment, string> = { Left: "left", Centered: "center", Right: "right", Justified: "both" };

function twips(pt: number): number {
  return Math.round(pt * 20);
}

function lineAttributes(setting: LineSpacingSetting): string {
  return setting.kind === "multiple"
    ? `w:line="${Math.round(setting.lines * 240)}" w:lineRule="auto"`
    : `w:line="${twips(setting.pt)}" w:lineRule="exact"`;
}

function toStyleXml(spec: StyleSpec): string {
  const { font, paragraph } = spec;
  const pPr =
    (paragraph.keepWithNext ? "<w:keepNext/>" : "") +
    `<w:spacing w:before="${twips(paragraph.spaceBeforePt)}" w:after="${twips(paragraph.spaceAfterPt)}" ${lineAttributes(paragraph.lineSpacing)}/>` +
    (paragraph.firstLineIndentPt > 0 ? `<w:ind w:firstLine="${twips(paragraph.firstLineIndentPt)}"/>` : "") +
    `<w:jc w:val="${JC[paragraph.alignment]}"/>`;
  const rPr =
    `<w:rFonts w:ascii="${font.name}" w:hAnsi="${font.name}" w:cs="${font.name}"/>` +
    (font.bold ? "<w:b/><w:bCs/>" : "") +
    (font.italic ? "<w:i/><w:iCs/>" : "") +
    `<w:color w:val="${font.color.replace("#", "")}"/>` +
    `<w:sz w:val="${font.sizePt * 2}"/><w:szCs w:val="${font.sizePt * 2}"/>` +
    (font.underline ? '<w:u w:val="single"/>' : "");
  return (
    `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${styleIdFromName(spec.name)}">` +
    `<w:name w:val="${spec.name}"/><w:basedOn w:val="Normal"/><w:qFormat/>` +
    `<w:pPr>${pPr}</w:pPr><w:rPr>${rPr}</w:rPr></w:style>`
  );
}

/**
 * Các phần tử w:style cho phần styles của gói OOXML. Là cách duy nhất đã biết đặt được giãn dòng
 * Exactly cho style, nhưng chỉ cài mới (không cập nhật style đã có) và phải chèn kèm một đoạn rỗng
 * phía sau (kết quả spike T3).
 */
export function toStyleDefinitionsXml(specs: readonly StyleSpec[]): string {
  return `${specs.map(toStyleXml).join("\n")}\n`;
}
