import { escapeXml } from "./headerBlock";

/**
 * Đánh số tự động Điều, Khoản, Điểm (T16): một danh sách nhiều cấp. Cấp 0 "Điều 1." (đậm), cấp 1 "1." (Khoản),
 * cấp 2 "a)" (Điểm); số Khoản bắt đầu lại ở mỗi Điều, số Điểm bắt đầu lại ở mỗi Khoản. Thêm hoặc xoá một Điều
 * thì Word tự đánh số lại. Chữ cái của Điểm theo bảng a-z của Word: Điểm "đ" (sau "d") phải gõ tay.
 */

export const NUMBERING_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** Mã danh sách dùng trong thân văn bản; Word cấp lại mã khi chèn nên chỉ cần khớp giữa phần numbering và thân. */
export const LEGAL_NUM_ID = 1;

export const LEGAL_LEVELS = { dieu: 0, khoan: 1, diem: 2 } as const;
export type LegalLevel = keyof typeof LEGAL_LEVELS;

function level(ilvl: number, fmt: "decimal" | "lowerLetter", text: string, bold: boolean, firstLineTwips: number, leftTwips: number): string {
  return (
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:suff w:val="space"/><w:lvlText w:val="${escapeXml(text)}"/><w:lvlJc w:val="left"/>` +
    `<w:pPr><w:ind w:left="${leftTwips}" w:firstLine="${firstLineTwips}"/></w:pPr><w:rPr>${bold ? "<w:b/>" : '<w:b w:val="0"/>'}</w:rPr></w:lvl>`
  );
}

/** Phần numbering.xml cho danh sách Điều, Khoản, Điểm; `firstLineTwips` là thụt dòng đầu của nội dung. */
export function legalNumberingXml(firstLineTwips: number): string {
  return (
    `<w:numbering xmlns:w="${NUMBERING_NS}">` +
    `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>` +
    level(0, "decimal", "Điều %1.", true, firstLineTwips, 0) +
    level(1, "decimal", "%2.", false, firstLineTwips, 0) +
    level(2, "lowerLetter", "%3)", false, firstLineTwips, 0) +
    `</w:abstractNum><w:num w:numId="${LEGAL_NUM_ID}"><w:abstractNumId w:val="0"/></w:num></w:numbering>`
  );
}

/** Đoạn đánh số tự động ở cấp `kind`, gắn style `styleId`; `rawXml` là phần ruột đã dựng (các run) hoặc để trống. */
export function legalParagraph(kind: LegalLevel, styleId: string, text: string, rawRuns = ""): string {
  const run = text === "" ? "" : `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
  return `<w:p><w:pPr><w:pStyle w:val="${styleId}"/><w:numPr><w:ilvl w:val="${LEGAL_LEVELS[kind]}"/><w:numId w:val="${LEGAL_NUM_ID}"/></w:numPr></w:pPr>${run}${rawRuns}</w:p>`;
}
