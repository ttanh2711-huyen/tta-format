import { BODY_MARKER } from "./documentAssembler";
import { escapeXml } from "./headerBlock";
import { multilevelNumberingXml, numberedParagraph, type LevelSpec } from "./numbering";
import { styleIdFromName } from "./styleXml";

/**
 * Trang phụ lục (T21): bắt đầu trang mới, nhãn "Phụ lục I" tự đánh số La Mã (bắt đầu từ số của phụ lục này nên các
 * phụ lục chèn ở các lần khác nhau nối tiếp nhau), tên phụ lục in hoa đậm, chỉ dẫn "(Kèm theo ...)" nghiêng. Dùng
 * style PhuLucSo, PhuLucTieuDe, PhuLucChiDan của chuẩn.
 */

export interface AppendixInput {
  /** Tên phụ lục, ví dụ "Danh sách đại biểu"; sẽ được in hoa. */
  title: string;
  /** Chỉ dẫn kèm theo, ví dụ "Kèm theo Công văn số 05/CV-VP ngày 05 tháng 01 năm 2026"; trống thì bỏ. */
  reference: string;
  /** Số thứ tự của phụ lục này, bắt đầu từ 1 (đếm các phụ lục đã có cộng 1). */
  number: number;
  stylePrefix: string;
  /** Có chừa chỗ nội dung (đoạn trống mang chữ đánh dấu) sau chỉ dẫn không. */
  withBody?: boolean;
}

export const MAX_APPENDIX = 99;

function clean(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

export function validateAppendix(input: AppendixInput): string[] {
  const errors: string[] = [];
  if (clean(input.title) === "") errors.push("Hãy nhập tên phụ lục.");
  if (!Number.isInteger(input.number) || input.number < 1 || input.number > MAX_APPENDIX) errors.push(`Số thứ tự phụ lục phải từ 1 đến ${MAX_APPENDIX}.`);
  return errors;
}

function levels(start: number): LevelSpec[] {
  return [{ id: "phuLuc", fmt: "upperRoman", text: "Phụ lục %1", bold: true, labelOnly: true, start }];
}

/** Phần numbering.xml của một phụ lục: số thứ tự La Mã bắt đầu từ `start`. */
export function appendixNumberingXml(start: number): string {
  return multilevelNumberingXml(levels(start));
}

/** Câu chỉ dẫn đặt trong ngoặc đơn nếu người dùng chưa gõ. */
export function referenceText(text: string): string {
  const value = clean(text);
  if (value === "") return "";
  return value.startsWith("(") ? value : `(${value})`;
}

/** XML thân văn bản của một trang phụ lục. */
export function buildAppendix(input: AppendixInput): string {
  const problems = validateAppendix(input);
  if (problems.length > 0) throw new Error(problems.join(" "));
  const id = (element: string) => styleIdFromName(`${input.stylePrefix}${element}`);
  // pageBreakBefore đứng giữa pStyle và numPr theo thứ tự của lược đồ.
  const label = numberedParagraph(levels(input.number), "phuLuc", id("PhuLucSo"), "").replace("<w:numPr>", "<w:pageBreakBefore/><w:numPr>");
  const parts = [label, `<w:p><w:pPr><w:pStyle w:val="${id("PhuLucTieuDe")}"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(clean(input.title).toUpperCase())}</w:t></w:r></w:p>`];
  const reference = referenceText(input.reference);
  if (reference !== "") parts.push(`<w:p><w:pPr><w:pStyle w:val="${id("PhuLucChiDan")}"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(reference)}</w:t></w:r></w:p>`);
  if (input.withBody) parts.push(`<w:p><w:pPr><w:pStyle w:val="${id("NoiDung")}"/></w:pPr><w:r><w:t>${BODY_MARKER}</w:t></w:r></w:p>`);
  return parts.join("");
}
