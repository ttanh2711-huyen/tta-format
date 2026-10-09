import { mmToTwips } from "./flatOpc";
import { escapeXml } from "./headerBlock";
import { textWidthMm } from "./textWidth";

/**
 * Khung dấu chỉ mức độ khẩn hoặc chỉ dẫn phạm vi lưu hành (T21): hình chữ nhật viền đơn, chữ in hoa đậm, canh giữa.
 * Dựng bằng bảng một ô có chiều cao cố định để Word không giãn khung theo chữ. Kích thước lấy từ quy định của dấu (nếu
 * chuẩn nêu); không nêu thì rộng vừa chữ cộng lề, cao 8 mm.
 */

const FONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>';
const DEFAULT_HEIGHT_MM = 8;
/** Lề hai bên chữ khi tự tính chiều rộng khung, mm. */
const PADDING_MM = 6;

export interface UrgencyMarkInput {
  /** Chữ trong khung; sẽ được in hoa. */
  text: string;
  widthMm?: number;
  heightMm?: number;
  sizePt: number;
  /** Căn khung trong vùng chữ. */
  align?: "left" | "center" | "right";
}

export function validateUrgencyMark(input: UrgencyMarkInput): string[] {
  const errors: string[] = [];
  if (input.text.trim() === "") errors.push("Hãy chọn hoặc nhập chữ trong khung.");
  if (input.widthMm !== undefined && !(input.widthMm > 0)) errors.push("Chiều rộng khung phải lớn hơn 0.");
  if (input.heightMm !== undefined && !(input.heightMm > 0)) errors.push("Chiều cao khung phải lớn hơn 0.");
  return errors;
}

/** Kích thước khung (mm) sẽ dùng. */
export function markSize(input: UrgencyMarkInput): { widthMm: number; heightMm: number } {
  const upper = input.text.replace(/\s+/g, " ").trim().toUpperCase();
  return { widthMm: input.widthMm ?? Math.ceil(textWidthMm(upper, input.sizePt, true)) + PADDING_MM, heightMm: input.heightMm ?? DEFAULT_HEIGHT_MM };
}

/** XML thân văn bản của khung (một bảng) và một đoạn đệm cao 1 pt phía sau để bảng không dính nội dung kế tiếp. */
export function buildUrgencyMark(input: UrgencyMarkInput): string {
  const problems = validateUrgencyMark(input);
  if (problems.length > 0) throw new Error(problems.join(" "));
  const { widthMm, heightMm } = markSize(input);
  const width = mmToTwips(widthMm);
  const height = mmToTwips(heightMm);
  const half = input.sizePt * 2;
  const text = escapeXml(input.text.replace(/\s+/g, " ").trim().toUpperCase());
  const border = ["top", "left", "bottom", "right"].map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join("");
  const jc = input.align && input.align !== "left" ? `<w:jc w:val="${input.align}"/>` : "";
  const run = `<w:rPr>${FONT}<w:b/><w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr>`;
  const tail = `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr>${FONT}<w:sz w:val="2"/></w:rPr></w:pPr></w:p>`;
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="${width}" w:type="dxa"/>${jc}<w:tblBorders>${border}</w:tblBorders><w:tblLayout w:type="fixed"/>` +
    `<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="${width}"/></w:tblGrid>` +
    `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="${height}" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>` +
    `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>${run}</w:pPr><w:r>${run}<w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>${tail}`
  );
}
