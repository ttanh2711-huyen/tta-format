import { mmToTwips } from "../../ooxml/flatOpc";
import { MAX_POSITION_LINES, recipientLines } from "../../ooxml/footerBlock";
import { borderlessTable, cell, escapeXml } from "../../ooxml/headerBlock";
import type { FormatValues } from "../../settings/formatOptions";
import type { StandardProfile } from "../types";

/**
 * Khối cuối văn bản Đảng (TD4): bảng 2 cột không viền. Trái là "Nơi nhận:" (đứng, gạch chân) và danh sách, có
 * thể kèm ký hiệu người soạn và số bản phát hành; phải là quyền hạn (đậm) cùng chức vụ (in hoa, không đậm), chỗ
 * ký, họ tên (đậm) và dòng xác nhận chữ ký tuỳ chọn.
 */

export { MAX_POSITION_LINES };
const MAX_SIGN_GAP = 12;
const LEFT_COLUMN_RATIO = 0.5;
const FONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>';

export interface PartyFooterInput {
  /** Các nơi nhận, không gồm mục "Lưu". */
  recipients: readonly string[];
  /** Nội dung sau "Lưu:", ví dụ "VT, TCTU"; trống thì không có mục Lưu. */
  luu: string;
  /** Ký hiệu người soạn thảo và số lượng bản phát hành (11 pt, cuối "Nơi nhận"); trống thì bỏ. */
  drafter: string;
  /** Quyền hạn ký (T/M, K/T, T/L, Q.); rỗng nếu ký trực tiếp. */
  authority: string;
  /** Chức vụ, một hoặc hai dòng; dòng đầu đi cùng quyền hạn. */
  positionLines: readonly string[];
  signerName: string;
  /** Họ tên và chức vụ người được xác nhận chữ ký ("Xác nhận chữ ký của đồng chí ..."); trống thì bỏ. */
  confirmation: string;
  signGapLines: number;
  sizes: { noiNhanTu: number; noiNhanDanhSach: number; quyenHan: number; chucVu: number; hoTen: number; kyHieuNguoiSoan: number };
  textWidthMm: number;
}

function clean(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

export function validatePartyFooter(input: PartyFooterInput): string[] {
  const errors: string[] = [];
  if (clean(input.signerName) === "") errors.push("Hãy nhập họ tên người ký.");
  if (input.positionLines.map(clean).filter((line) => line !== "").length === 0) errors.push("Hãy nhập chức vụ người ký.");
  if (input.positionLines.length > MAX_POSITION_LINES) errors.push(`Chức vụ tối đa ${MAX_POSITION_LINES} dòng.`);
  if (!Number.isInteger(input.signGapLines) || input.signGapLines < 1 || input.signGapLines > MAX_SIGN_GAP) {
    errors.push(`Số dòng trống chỗ ký phải từ 1 đến ${MAX_SIGN_GAP}.`);
  }
  return errors;
}

interface Run {
  text: string;
  bold?: boolean;
  underline?: boolean;
}

/** Đoạn gồm nhiều run (có thể khác đậm, gạch chân); không có run nào thì là đoạn trống mang cỡ chữ. */
function paragraph(runs: readonly Run[], sizePt: number, align: "left" | "center", beforePt = 0): string {
  const half = sizePt * 2;
  const props = (run: Run) => `${FONT}${run.bold ? "<w:b/>" : ""}${run.underline ? '<w:u w:val="single"/>' : ""}<w:sz w:val="${half}"/><w:szCs w:val="${half}"/>`;
  const body = runs.map((run) => `<w:r><w:rPr>${props(run)}</w:rPr><w:t xml:space="preserve">${escapeXml(run.text)}</w:t></w:r>`).join("");
  const mark = `<w:rPr>${FONT}<w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr>`;
  return `<w:p><w:pPr><w:spacing w:before="${beforePt * 20}" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="${align}"/>${mark}</w:pPr>${body}</w:p>`;
}

/** XML thân văn bản của khối cuối: bảng và một đoạn đệm cao 1 pt phía sau để không đẩy thêm trang trắng. */
export function buildPartyFooterBlock(input: PartyFooterInput): string {
  const problems = validatePartyFooter(input);
  if (problems.length > 0) throw new Error(problems.join(" "));

  const total = mmToTwips(input.textWidthMm);
  const leftWidth = Math.round(total * LEFT_COLUMN_RATIO);
  const rightWidth = total - leftWidth;
  const { sizes } = input;

  const lines = recipientLines(input.recipients, input.luu);
  const drafter = clean(input.drafter);
  const left =
    lines.length === 0
      ? [paragraph([], sizes.noiNhanDanhSach, "left")]
      : [
          paragraph([{ text: "Nơi nhận:", underline: true }], sizes.noiNhanTu, "left"),
          ...lines.map((line) => paragraph([{ text: line }], sizes.noiNhanDanhSach, "left")),
          ...(drafter === "" ? [] : [paragraph([{ text: drafter }], sizes.kyHieuNguoiSoan, "left", 6)]),
        ];

  const positions = input.positionLines.map(clean).filter((line) => line !== "").map((line) => line.toUpperCase());
  const authority = clean(input.authority);
  const first = authority === "" ? [{ text: positions[0] }] : [{ text: authority, bold: true }, { text: ` ${positions[0]}` }];
  const confirmation = clean(input.confirmation);
  const right = [
    paragraph(first, sizes.chucVu, "center"),
    ...positions.slice(1).map((line) => paragraph([{ text: line }], sizes.chucVu, "center")),
    ...Array.from({ length: input.signGapLines }, () => paragraph([], sizes.chucVu, "center")),
    paragraph([{ text: clean(input.signerName), bold: true }], sizes.hoTen, "center"),
    ...(confirmation === "" ? [] : [paragraph([{ text: `Xác nhận chữ ký của đồng chí ${confirmation}` }], sizes.chucVu, "center", 6)]),
  ];

  const tail = `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr>${FONT}<w:sz w:val="2"/></w:rPr></w:pPr></w:p>`;
  return borderlessTable(total, leftWidth, `<w:tr>${cell(leftWidth, left)}${cell(rightWidth, right)}</w:tr>`) + tail;
}

/** Dựng thông số khối cuối chuẩn Đảng từ Tuỳ chọn nhanh: cỡ chữ, số dòng trống chỗ ký, bề rộng vùng chữ. */
export function partyFooterFromOptions(
  profile: StandardProfile,
  values: FormatValues,
  text: Pick<PartyFooterInput, "recipients" | "luu" | "drafter" | "authority" | "positionLines" | "signerName" | "confirmation">,
): PartyFooterInput {
  const num = (id: string, fallback: number): number => (typeof values[id] === "number" ? (values[id] as number) : fallback);
  const fixed = profile.pageDefaults.marginsMm;
  return {
    ...text,
    signGapLines: num("signGap", 6),
    sizes: {
      noiNhanTu: num("size.noiNhanTu", 14),
      noiNhanDanhSach: num("size.noiNhanDanhSach", 12),
      quyenHan: num("size.quyenHan", 14),
      chucVu: num("size.chucVu", 14),
      hoTen: num("size.hoTen", 14),
      kyHieuNguoiSoan: num("size.kyHieuNguoiSoan", 11),
    },
    textWidthMm: profile.page.paper.widthMm - num("marginLeft", fixed.left) - num("marginRight", fixed.right),
  };
}
