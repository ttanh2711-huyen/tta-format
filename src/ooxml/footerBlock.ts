import type { FormatValues } from "../settings/formatOptions";
import type { StandardProfile } from "../standards/types";
import { mmToTwips } from "./flatOpc";

/**
 * Khối cuối văn bản chuẩn Chính quyền (T10): bảng 2 cột không viền, trái là "Nơi nhận" và danh sách,
 * phải là quyền hạn, chức vụ, chỗ ký và họ tên người ký. Theo cấu trúc của file mẫu công văn.
 */

export const MAX_POSITION_LINES = 2;
const MAX_SIGN_GAP = 12;
/** Tỉ lệ bề rộng cột trái (4660/9060 trong file mẫu). */
const LEFT_COLUMN_RATIO = 4660 / 9060;
const FONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>';

export interface FooterBlockInput {
  /** Các nơi nhận, không gồm mục "Lưu". */
  recipients: readonly string[];
  /** Nội dung sau "Lưu:", ví dụ "VT, CCTLTTA"; trống thì không có mục Lưu. */
  luu: string;
  /** Quyền hạn ký (TM., KT., TL., TUQ., Q.); chuỗi rỗng nếu ký trực tiếp. */
  authority: string;
  /** Chức vụ, một hoặc hai dòng; dòng đầu đi cùng quyền hạn. */
  positionLines: readonly string[];
  /**
   * Chức danh thứ hai của người ký có hai chức danh (ví dụ "TRƯỞNG BAN" ở trên, "THỦ TƯỚNG" ở dưới): in hoa, đậm, đặt
   * dưới chỗ ký, ngay trên họ tên. Trống thì không có.
   */
  secondTitle?: string;
  signerName: string;
  /** Số dòng trống chừa chỗ ký. */
  signGapLines: number;
  sizes: { noiNhanTieuDe: number; noiNhanDanhSach: number; quyenHanChucVu: number; hoTen: number };
  textWidthMm: number;
}

function clean(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

function escapeXml(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Danh sách nơi nhận hoàn chỉnh: gạch đầu dòng, ";" cuối mục, "." cuối mục cuối; mục "Lưu" luôn ở cuối. */
export function recipientLines(items: readonly string[], luu: string): string[] {
  const strip = (text: string): string => clean(text).replace(/^[-–]\s*/, "").replace(/[;.,]+$/, "").trim();
  const entries = items.map(strip).filter((item) => item !== "");
  const luuText = strip(luu);
  if (luuText !== "") entries.push(/^lưu\s*:/i.test(luuText) ? `Lưu:${luuText.replace(/^lưu\s*:/i, "").replace(/^\s*/, " ")}` : `Lưu: ${luuText}`);
  return entries.map((entry, index) => `- ${entry}${index === entries.length - 1 ? "." : ";"}`);
}

const ACADEMIC_ABBREVIATIONS = /^(?:PGS|GS|TSKH|TS|Th\.S|ThS|NCS|CN|KS|BSCKII|BSCKI|BSNT|BS|CKII|CKI|DS)(?=[.\s]|$)/i;
const ACADEMIC_WORDS = /^(?:phó giáo sư|giáo sư|tiến sĩ|thạc sĩ|bác sĩ|dược sĩ|kỹ sư|cử nhân)(?=\s|$)/i;

/** Học hàm, học vị đứng đầu họ tên (NĐ 30 không ghi học hàm, học vị ở chữ ký); trả về cụm tìm thấy hoặc null. */
export function findAcademicTitle(name: string): string | null {
  const text = clean(name);
  return ACADEMIC_ABBREVIATIONS.exec(text)?.[0] ?? ACADEMIC_WORDS.exec(text)?.[0] ?? null;
}

/** Bỏ học hàm, học vị đứng đầu họ tên ("PGS.TS Lê Thị B" thành "Lê Thị B"). */
export function stripAcademicTitles(name: string): string {
  let rest = clean(name);
  for (let guard = 0; guard < 6; guard += 1) {
    const found = findAcademicTitle(rest);
    if (found === null) break;
    rest = rest.slice(found.length).replace(/^[.\s]+/, "");
  }
  return rest;
}

/** Chức danh thứ hai dưới chỗ ký, in hoa; trống thì không có dòng nào. */
export function secondTitleLines(secondTitle: string | undefined): string[] {
  const text = clean(secondTitle ?? "");
  return text === "" ? [] : [text.toUpperCase()];
}

export function validateFooterInput(input: FooterBlockInput): string[] {
  const errors: string[] = [];
  if (clean(input.signerName) === "") errors.push("Hãy nhập họ tên người ký.");
  const lines = input.positionLines.map(clean).filter((line) => line !== "");
  if (lines.length === 0) errors.push("Hãy nhập chức vụ người ký.");
  if (input.positionLines.length > MAX_POSITION_LINES) errors.push(`Chức vụ tối đa ${MAX_POSITION_LINES} dòng.`);
  if (!Number.isInteger(input.signGapLines) || input.signGapLines < 1 || input.signGapLines > MAX_SIGN_GAP) {
    errors.push(`Số dòng trống chỗ ký phải từ 1 đến ${MAX_SIGN_GAP}.`);
  }
  return errors;
}

function paragraph(text: string, sizePt: number, style: { bold?: boolean; italic?: boolean; align: "left" | "center" }): string {
  const half = sizePt * 2;
  const run = text === "" ? "" : `<w:r><w:rPr>${FONT}${style.bold ? "<w:b/>" : ""}${style.italic ? "<w:i/>" : ""}<w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
  const mark = `<w:rPr>${FONT}${style.bold ? "<w:b/>" : ""}${style.italic ? "<w:i/>" : ""}<w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr>`;
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="${style.align}"/>${mark}</w:pPr>${run}</w:p>`;
}

function cell(widthTwips: number, paragraphs: string[]): string {
  return `<w:tc><w:tcPr><w:tcW w:w="${widthTwips}" w:type="dxa"/></w:tcPr>${paragraphs.join("")}</w:tc>`;
}

const NO_BORDERS = ["top", "left", "bottom", "right", "insideH", "insideV"].map((side) => `<w:${side} w:val="none" w:sz="0" w:space="0" w:color="auto"/>`).join("");

/**
 * XML thân văn bản của khối cuối: bảng và một đoạn đệm cao 1 pt phía sau (như file mẫu) để không đẩy
 * thêm trang trắng.
 */
export function buildFooterBlock(input: FooterBlockInput): string {
  const problems = validateFooterInput(input);
  if (problems.length > 0) throw new Error(problems.join(" "));

  const total = mmToTwips(input.textWidthMm);
  const leftWidth = Math.round(total * LEFT_COLUMN_RATIO);
  const rightWidth = total - leftWidth;
  const { sizes } = input;

  const lines = recipientLines(input.recipients, input.luu);
  const left =
    lines.length === 0
      ? [paragraph("", sizes.noiNhanDanhSach, { align: "left" })]
      : [
          paragraph("Nơi nhận:", sizes.noiNhanTieuDe, { bold: true, italic: true, align: "left" }),
          ...lines.map((line) => paragraph(line, sizes.noiNhanDanhSach, { align: "left" })),
        ];

  const positions = input.positionLines.map(clean).filter((line) => line !== "").map((line) => line.toUpperCase());
  const authority = clean(input.authority);
  positions[0] = authority === "" ? positions[0] : `${authority} ${positions[0]}`;
  const right = [
    ...positions.map((line) => paragraph(line, sizes.quyenHanChucVu, { bold: true, align: "center" })),
    ...Array.from({ length: input.signGapLines }, () => paragraph("", sizes.quyenHanChucVu, { bold: true, align: "center" })),
    ...secondTitleLines(input.secondTitle).map((line) => paragraph(line, sizes.quyenHanChucVu, { bold: true, align: "center" })),
    paragraph(clean(input.signerName), sizes.hoTen, { bold: true, align: "center" }),
  ];

  const tail = `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr>${FONT}<w:sz w:val="2"/></w:rPr></w:pPr></w:p>`;
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="${total}" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>${NO_BORDERS}</w:tblBorders>` +
    `<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="${leftWidth}"/><w:gridCol w:w="${rightWidth}"/></w:tblGrid>` +
    `<w:tr>${cell(leftWidth, left)}${cell(rightWidth, right)}</w:tr></w:tbl>${tail}`
  );
}

/** Dựng thông số khối cuối từ Tuỳ chọn nhanh: cỡ chữ, số dòng trống chỗ ký, bề rộng vùng chữ. */
export function footerInputFromOptions(
  profile: StandardProfile,
  values: FormatValues,
  text: { recipients: readonly string[]; luu: string; authority: string; positionLines: readonly string[]; secondTitle?: string; signerName: string },
): FooterBlockInput {
  const num = (id: string, fallback: number): number => (typeof values[id] === "number" ? (values[id] as number) : fallback);
  const fixed = profile.pageDefaults.marginsMm;
  return {
    ...text,
    signGapLines: num("signGap", 6),
    sizes: {
      noiNhanTieuDe: num("size.noiNhanTieuDe", 12),
      noiNhanDanhSach: num("size.noiNhanDanhSach", 11),
      quyenHanChucVu: num("size.quyenHanChucVu", 14),
      hoTen: num("size.hoTenNguoiKy", 14),
    },
    textWidthMm: profile.page.paper.widthMm - num("marginLeft", fixed.left) - num("marginRight", fixed.right),
  };
}
