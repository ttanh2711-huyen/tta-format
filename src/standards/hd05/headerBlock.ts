import { mmToTwips } from "../../ooxml/flatOpc";
import { borderlessTable, cell, LEFT_COLUMN_RATIO, lineParagraph, textParagraph } from "../../ooxml/headerBlock";
import { textWidthMm } from "../../ooxml/textWidth";
import type { FormatValues } from "../../settings/formatOptions";
import type { StandardProfile } from "../types";
import { HEADER_MARKS } from "./rules";

/**
 * Khối đầu văn bản Đảng (TD1): bảng 2 cột không viền. Trái là cơ quan cấp trên trực tiếp (in hoa, không
 * đậm), cơ quan ban hành (in hoa, đậm, nhiều dòng được) và dấu * ngay dưới; phải là tiêu đề "ĐẢNG CỘNG SẢN
 * VIỆT NAM" (in hoa, đậm) với đường kẻ dài bằng tiêu đề. Hàng số, ký hiệu và ngày tháng thêm ở TD2.
 */

export const MAX_PARENT_LINES = 2;
export const MAX_ISSUER_LINES = 3;

export interface PartyHeaderInput {
  /** Cơ quan cấp trên trực tiếp, từ 0 đến 2 dòng; dòng trống bị bỏ. */
  parentLines: readonly string[];
  /** Cơ quan ban hành, từ 1 đến 3 dòng (liên cơ quan nhập mỗi cơ quan một dòng hoặc nối bằng dấu gạch nối). */
  issuerLines: readonly string[];
  sizes: { title: number; parent: number; issuer: number; so: number; placeDate: number };
  /** Hàng thứ hai: số và ký hiệu (trái, dưới tên cơ quan) và địa danh, ngày tháng (phải), đã dựng sẵn thành chữ (TD2). */
  numberRow?: { so: string; placeDate: string };
  /** Hàng thứ ba của công văn: dòng "về việc ..." (nghiêng) ở ô trái, ô phải để trống (TD3). */
  subjectRow?: { text: string; sizePt: number };
  /** Bề rộng vùng chữ (mm) mà bảng phải vừa. */
  textWidthMm: number;
}

function clean(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

export function validatePartyHeader(input: PartyHeaderInput): string[] {
  const errors: string[] = [];
  const issuers = input.issuerLines.map(clean).filter((line) => line !== "");
  if (issuers.length === 0) errors.push("Hãy nhập tên cơ quan ban hành.");
  if (input.issuerLines.length > MAX_ISSUER_LINES) errors.push(`Cơ quan ban hành tối đa ${MAX_ISSUER_LINES} dòng.`);
  if (input.parentLines.length > MAX_PARENT_LINES) errors.push(`Cơ quan cấp trên tối đa ${MAX_PARENT_LINES} dòng.`);
  return errors;
}

/** XML thân văn bản (một bảng) của khối đầu chuẩn Đảng; người gọi thêm đoạn phía sau nếu cần. */
export function buildPartyHeaderBlock(input: PartyHeaderInput): string {
  const problems = validatePartyHeader(input);
  if (problems.length > 0) throw new Error(problems.join(" "));

  const total = mmToTwips(input.textWidthMm);
  const leftWidth = Math.round(total * LEFT_COLUMN_RATIO);
  const rightWidth = total - leftWidth;
  const rightMm = (rightWidth * 25.4) / 1440;
  const upper = (lines: readonly string[]) => lines.map(clean).filter((line) => line !== "").map((line) => line.toUpperCase());

  const title = HEADER_MARKS.identifyingTitle;
  const left = [
    ...upper(input.parentLines).map((line) => textParagraph(line, input.sizes.parent)),
    ...upper(input.issuerLines).map((line) => textParagraph(line, input.sizes.issuer, { bold: true })),
    textParagraph(HEADER_MARKS.agencyStar, input.sizes.issuer),
  ];
  const right = [textParagraph(title, input.sizes.title, { bold: true }), lineParagraph(Math.min(textWidthMm(title, input.sizes.title, true), rightMm), rightWidth)];

  const numberRow = input.numberRow
    ? `<w:tr>${cell(leftWidth, [textParagraph(input.numberRow.so, input.sizes.so, { beforePt: 6 })])}${cell(rightWidth, [
        textParagraph(input.numberRow.placeDate, input.sizes.placeDate, { italic: true, beforePt: 6 }),
      ])}</w:tr>`
    : "";
  const subjectRow = input.subjectRow
    ? `<w:tr>${cell(leftWidth, [textParagraph(input.subjectRow.text, input.subjectRow.sizePt, { italic: true, beforePt: 6 })])}${cell(rightWidth, [textParagraph("", input.subjectRow.sizePt)])}</w:tr>`
    : "";
  return borderlessTable(total, leftWidth, `<w:tr>${cell(leftWidth, left)}${cell(rightWidth, right)}</w:tr>${numberRow}${subjectRow}`);
}

/** Dựng thông số khối đầu chuẩn Đảng từ Tuỳ chọn nhanh: cỡ chữ và bề rộng vùng chữ theo lề đang chọn. */
export function partyHeaderFromOptions(
  profile: StandardProfile,
  values: FormatValues,
  text: { parentLines: readonly string[]; issuerLines: readonly string[]; numberRow?: { so: string; placeDate: string }; subjectRow?: { text: string; sizePt: number } },
): PartyHeaderInput {
  const num = (id: string, fallback: number): number => (typeof values[id] === "number" ? (values[id] as number) : fallback);
  const fixed = profile.pageDefaults.marginsMm;
  return {
    ...text,
    sizes: {
      title: num("size.tieuDe", 15),
      parent: num("size.coQuanCapTren", 14),
      issuer: num("size.coQuanBanHanh", 14),
      so: num("size.soKyHieu", 14),
      placeDate: num("size.diaDanhNgayThang", 14),
    },
    textWidthMm: profile.page.paper.widthMm - num("marginLeft", fixed.left) - num("marginRight", fixed.right),
  };
}
