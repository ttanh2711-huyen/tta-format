import { escapeXml } from "./headerBlock";

/**
 * Đánh số tự động nhiều cấp cho Phần/Chương/Mục/Điều/Khoản/Điểm (T16, T18, TD6). Mỗi bộ cấp (`LevelSpec[]`) thành một
 * danh sách nhiều cấp trong numbering.xml; thêm hoặc xoá một mục thì Word tự đánh số lại. Chữ cái của Điểm theo bảng
 * a-z của Word: Điểm "đ" (sau "d") phải gõ tay.
 */

export const NUMBERING_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** Mã danh sách dùng trong thân văn bản; Word cấp lại mã khi chèn nên chỉ cần khớp giữa phần numbering và thân. */
export const LEGAL_NUM_ID = 1;

export type NumberFormat = "decimal" | "lowerLetter" | "upperRoman" | "vietnameseCounting";

export interface LevelSpec {
  /** Tên cấp, dùng để gọi `numberedParagraph`. */
  id: string;
  fmt: NumberFormat;
  /** Mẫu nhãn, ví dụ "Điều %1." (số thứ tự cấp k ghi là `%k+1`). */
  text: string;
  bold: boolean;
  /** Nhãn đứng một mình (đoạn không có chữ); không thì cách chữ một khoảng trắng. */
  labelOnly?: boolean;
  /** Không bắt đầu lại khi cấp trên tăng (số Điều chạy liên tục qua các Chương, Mục). */
  continuous?: boolean;
  /** Thụt dòng đầu (twips) của đoạn mang nhãn. */
  firstLineTwips?: number;
  /** Số bắt đầu của cấp (mặc định 1); dùng để nối số thứ tự giữa các lần chèn. */
  start?: number;
}

function level(index: number, spec: LevelSpec): string {
  const suff = spec.labelOnly ? "nothing" : "space";
  const restart = spec.continuous ? '<w:lvlRestart w:val="0"/>' : "";
  return (
    `<w:lvl w:ilvl="${index}"><w:start w:val="${spec.start ?? 1}"/><w:numFmt w:val="${spec.fmt}"/>${restart}<w:suff w:val="${suff}"/><w:lvlText w:val="${escapeXml(spec.text)}"/><w:lvlJc w:val="left"/>` +
    `<w:pPr><w:ind w:left="0" w:firstLine="${spec.firstLineTwips ?? 0}"/></w:pPr><w:rPr>${spec.bold ? "<w:b/>" : '<w:b w:val="0"/>'}</w:rPr></w:lvl>`
  );
}

/** Phần numbering.xml cho một danh sách nhiều cấp (cả thẻ `w:numbering`). */
export function multilevelNumberingXml(levels: readonly LevelSpec[]): string {
  return (
    `<w:numbering xmlns:w="${NUMBERING_NS}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>` +
    levels.map((spec, index) => level(index, spec)).join("") +
    `</w:abstractNum><w:num w:numId="${LEGAL_NUM_ID}"><w:abstractNumId w:val="0"/></w:num></w:numbering>`
  );
}

/** Đoạn đánh số tự động ở cấp `levelId` của `levels`, gắn style `styleId`; `rawRuns` là các run đã dựng (ví dụ chỗ đánh dấu). */
export function numberedParagraph(levels: readonly LevelSpec[], levelId: string, styleId: string, text: string, rawRuns = ""): string {
  const index = levels.findIndex((spec) => spec.id === levelId);
  if (index < 0) throw new Error(`Không có cấp "${levelId}".`);
  const run = text === "" ? "" : `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
  return `<w:p><w:pPr><w:pStyle w:val="${styleId}"/><w:numPr><w:ilvl w:val="${index}"/><w:numId w:val="${LEGAL_NUM_ID}"/></w:numPr></w:pPr>${run}${rawRuns}</w:p>`;
}

/** Bộ cấp của quyết định, nghị quyết (T16): Điều (đậm), Khoản, Điểm. */
export function legalLevels(firstLineTwips: number): LevelSpec[] {
  return [
    { id: "dieu", fmt: "decimal", text: "Điều %1.", bold: true, firstLineTwips },
    { id: "khoan", fmt: "decimal", text: "%2.", bold: false, firstLineTwips },
    { id: "diem", fmt: "lowerLetter", text: "%3)", bold: false, firstLineTwips },
  ];
}

export const LEGAL_LEVELS = { dieu: 0, khoan: 1, diem: 2 } as const;
export type LegalLevel = keyof typeof LEGAL_LEVELS;

/** Phần numbering.xml cho danh sách Điều, Khoản, Điểm; `firstLineTwips` là thụt dòng đầu của nội dung. */
export function legalNumberingXml(firstLineTwips: number): string {
  return multilevelNumberingXml(legalLevels(firstLineTwips));
}

/** Đoạn đánh số tự động ở cấp `kind` của danh sách Điều, Khoản, Điểm. */
export function legalParagraph(kind: LegalLevel, styleId: string, text: string, rawRuns = ""): string {
  return numberedParagraph(legalLevels(0), kind, styleId, text, rawRuns);
}

/** Bộ cấp của quy chế, quy định (T18): Chương, Mục, Điều (chạy liên tục), Khoản, Điểm. */
export function regulationLevels(firstLineTwips: number): LevelSpec[] {
  return [
    { id: "chuong", fmt: "upperRoman", text: "Chương %1", bold: true, labelOnly: true },
    { id: "muc", fmt: "decimal", text: "Mục %2", bold: true, labelOnly: true },
    { id: "dieu", fmt: "decimal", text: "Điều %3.", bold: true, continuous: true, firstLineTwips },
    { id: "khoan", fmt: "decimal", text: "%4.", bold: false, firstLineTwips },
    { id: "diem", fmt: "lowerLetter", text: "%5)", bold: false, firstLineTwips },
  ];
}
