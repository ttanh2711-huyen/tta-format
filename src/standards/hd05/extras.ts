import { paragraph } from "../../ooxml/titleBlock";
import { multilevelNumberingXml, numberedParagraph, type LevelSpec } from "../../ooxml/numbering";
import { escapeXml } from "../../ooxml/headerBlock";
import { buildUrgencyMark } from "../../ooxml/urgencyMark";
import { paragraphIssue } from "../../checker/common";
import type { CheckDocument, Issue } from "../../checker/types";
import type { StandardProfile } from "../types";

/**
 * Thành phần bổ sung của văn bản Đảng (TD8, mục II của HD 05): chỉ dẫn phạm vi lưu hành (khung, 12 pt, in hoa đậm, góc
 * phải), chỉ dẫn dự thảo (14 pt nghiêng, dưới số và ký hiệu), thông tin liên hệ (10 pt, trang cuối, dưới đường kẻ ngang),
 * dòng "Không phổ biến trên phương tiện thông tin đại chúng" và trang phụ lục. Dấu chỉ mức độ khẩn dùng chung với
 * chuẩn Chính quyền (`ooxml/urgencyMark.ts`).
 */

export const NO_PUBLIC_NOTE = "Không phổ biến trên phương tiện thông tin đại chúng";
const DRAFT_PATTERN = /^Dự thảo(\s+lần\s+\d+)?$/i;

function clean(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

/** Khung chỉ dẫn phạm vi lưu hành: hộp viền đơn, in hoa đậm, canh phải, cỡ chữ theo thành phần của chuẩn. */
export function buildCirculationMark(text: string, sizePt: number): string {
  return buildUrgencyMark({ text, sizePt, align: "right" });
}

/** Chỉ dẫn dự thảo, ví dụ "Dự thảo lần 2": nghiêng, canh trái, cỡ 14. */
export function buildDraftNote(text: string, sizePt: number): string {
  const value = clean(text);
  if (!DRAFT_PATTERN.test(value)) throw new Error('Chỉ dẫn dự thảo có dạng "Dự thảo" hoặc "Dự thảo lần 2".');
  return paragraph(value, { sizePt, italic: true, align: "left" });
}

/** Dòng "Không phổ biến ..." ở cuối trang cuối: in hoa không, cỡ 12, canh giữa. */
export function buildNoPublicNote(sizePt: number): string {
  return paragraph(NO_PUBLIC_NOTE, { sizePt, align: "center", beforePt: 6 });
}

const FONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>';

/** Thông tin liên hệ: đường kẻ ngang hết vùng trình bày rồi các dòng chữ 10 pt (địa chỉ, điện thoại, thư điện tử). */
export function buildContactInfo(lines: readonly string[], sizePt: number): string {
  const entries = lines.map(clean).filter((line) => line !== "");
  if (entries.length === 0) throw new Error("Hãy nhập thông tin liên hệ.");
  const half = sizePt * 2;
  const rule = `<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr><w:spacing w:before="120" w:after="0" w:line="120" w:lineRule="exact"/></w:pPr></w:p>`;
  const rows = entries.map((line) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/></w:pPr><w:r><w:rPr>${FONT}<w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`);
  return rule + rows.join("");
}

/** Trang phụ lục của Đảng: "Phụ lục I" (nhãn tự đánh số), tên phụ lục in hoa đậm, chỉ dẫn nghiêng; bắt đầu trang mới. */
export interface PartyAppendixInput {
  title: string;
  reference: string;
  number: number;
  sizePt: number;
}

function levels(start: number): LevelSpec[] {
  return [{ id: "phuLuc", fmt: "upperRoman", text: "Phụ lục %1", bold: true, labelOnly: true, start }];
}

export function partyAppendixNumberingXml(start: number): string {
  return multilevelNumberingXml(levels(start));
}

export function validatePartyAppendix(input: PartyAppendixInput): string[] {
  const errors: string[] = [];
  if (clean(input.title) === "") errors.push("Hãy nhập tên phụ lục.");
  if (!Number.isInteger(input.number) || input.number < 1 || input.number > 99) errors.push("Số thứ tự phụ lục phải từ 1 đến 99.");
  return errors;
}

export function buildPartyAppendix(input: PartyAppendixInput): string {
  const problems = validatePartyAppendix(input);
  if (problems.length > 0) throw new Error(problems.join(" "));
  const half = input.sizePt * 2;
  const label = numberedParagraph(levels(input.number), "phuLuc", "Normal", "")
    .replace('<w:pStyle w:val="Normal"/>', "")
    .replace("<w:numPr>", "<w:pageBreakBefore/><w:numPr>")
    .replace("</w:numPr></w:pPr>", `</w:numPr><w:jc w:val="center"/><w:rPr><w:b/><w:sz w:val="${half}"/></w:rPr></w:pPr>`);
  const parts = [label, paragraph(clean(input.title).toUpperCase(), { sizePt: input.sizePt, bold: true })];
  const reference = clean(input.reference);
  if (reference !== "") parts.push(paragraph(reference.startsWith("(") ? reference : `(${reference})`, { sizePt: input.sizePt, italic: true, afterPt: 6 }));
  return parts.join("");
}

/** Nhận biết chỉ dẫn dự thảo và dòng không phổ biến: chỉ dẫn dự thảo nghiêng, cỡ 14; dòng không phổ biến cỡ 12. */
export function checkPartyExtras(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const draft = profile.elements.duThao;
  const circulation = profile.elements.phamViLuuHanh;
  if (!draft || !circulation) return [];
  const issues: Issue[] = [];
  for (const p of doc.paragraphs) {
    const text = p.resolved.text.trim();
    if (DRAFT_PATTERN.test(text)) {
      if (p.resolved.italic !== true) issues.push(paragraphIssue(p, { ruleId: "extras.draft.italic", severity: "error", message: "Chỉ dẫn dự thảo phải in nghiêng.", reference: draft.source, fix: { kind: "font", index: p.index, italic: true, label: "In nghiêng chỉ dẫn dự thảo" } }));
      const size = p.resolved.sizePt;
      if (size !== undefined && (size < draft.sizePt.min || size > draft.sizePt.max)) {
        issues.push(paragraphIssue(p, { ruleId: "extras.draft.size", severity: "error", message: `Chỉ dẫn dự thảo cỡ ${size} pt; phải ${draft.sizePt.min} pt.`, reference: draft.source, fix: { kind: "font", index: p.index, sizePt: draft.sizePt.min, label: `Đặt cỡ chữ ${draft.sizePt.min} pt` } }));
      }
    } else if (text === NO_PUBLIC_NOTE) {
      const size = p.resolved.sizePt;
      if (size !== undefined && (size < circulation.sizePt.min || size > circulation.sizePt.max)) {
        issues.push(paragraphIssue(p, { ruleId: "extras.noPublic.size", severity: "error", message: `Dòng "${NO_PUBLIC_NOTE}" cỡ ${size} pt; phải ${circulation.sizePt.min} pt.`, reference: circulation.source, fix: { kind: "font", index: p.index, sizePt: circulation.sizePt.min, label: `Đặt cỡ chữ ${circulation.sizePt.min} pt` } }));
      }
    }
  }
  return issues;
}
