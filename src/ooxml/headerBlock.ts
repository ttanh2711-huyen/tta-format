import type { FormatValues } from "../settings/formatOptions";
import type { StandardProfile } from "../standards/types";
import { mmToTwips } from "./flatOpc";
import { textWidthMm } from "./textWidth";

/**
 * Khối đầu văn bản chuẩn Chính quyền (T7): bảng 2 cột không viền, trái là cơ quan chủ quản và cơ quan ban
 * hành, phải là quốc hiệu và tiêu ngữ; mỗi bên có đường kẻ dưới bằng viền của một đoạn riêng (không dùng hình
 * nổi nên đi theo văn bản). Hàng số/ký hiệu, địa danh và V/v thêm ở T8, T9.
 */

export const QUOC_HIEU = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
export const TIEU_NGU = "Độc lập - Tự do - Hạnh phúc";
export const MAX_PARENT_LINES = 3;

/** Tỉ lệ bề rộng cột trái của bảng đầu văn bản (theo file mẫu). */
const LEFT_COLUMN_RATIO = 0.45;
const FONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>';

export interface HeaderBlockInput {
  /** Cơ quan chủ quản, từ 0 đến 3 dòng; dòng trống bị bỏ. */
  parentLines: readonly string[];
  issuer: string;
  sizes: { parent: number; issuer: number; quocHieu: number; tieuNgu: number; so: number; placeDate: number };
  /** Hàng thứ hai: dòng số, ký hiệu (trái) và địa danh, ngày tháng (phải), đã dựng sẵn thành chữ (T8). */
  numberRow?: { so: string; placeDate: string };
  /** Độ dài đường kẻ dưới tên cơ quan so với độ dài dòng chữ. */
  agencyLineRatio: number;
  /** Bề rộng vùng chữ (mm) mà bảng phải vừa. */
  textWidthMm: number;
}

function clean(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

function escapeXml(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function validateHeaderInput(input: HeaderBlockInput): string[] {
  const errors: string[] = [];
  if (clean(input.issuer) === "") errors.push("Hãy nhập tên cơ quan ban hành.");
  if (input.parentLines.length > MAX_PARENT_LINES) errors.push(`Cơ quan chủ quản tối đa ${MAX_PARENT_LINES} dòng.`);
  return errors;
}

function textParagraph(text: string, sizePt: number, style: { bold?: boolean; italic?: boolean; beforePt?: number } = {}): string {
  const half = sizePt * 2;
  const before = (style.beforePt ?? 0) * 20;
  return (
    `<w:p><w:pPr><w:spacing w:before="${before}" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>` +
    `<w:r><w:rPr>${FONT}${style.bold ? "<w:b/>" : ""}${style.italic ? "<w:i/>" : ""}<w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`
  );
}

/** Đoạn rỗng cao 6 pt có viền dưới, thụt hai bên để đường kẻ dài `lineMm` và nằm giữa ô rộng `cellTwips`. */
function lineParagraph(lineMm: number, cellTwips: number): string {
  const inset = Math.max(0, Math.round((cellTwips - mmToTwips(lineMm)) / 2));
  return (
    `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>` +
    `<w:spacing w:before="0" w:after="0" w:line="120" w:lineRule="exact"/>` +
    `<w:ind w:left="${inset}" w:right="${inset}"/><w:jc w:val="center"/></w:pPr></w:p>`
  );
}

function cell(widthTwips: number, paragraphs: string[]): string {
  return `<w:tc><w:tcPr><w:tcW w:w="${widthTwips}" w:type="dxa"/></w:tcPr>${paragraphs.join("")}</w:tc>`;
}

const NO_BORDERS = ["top", "left", "bottom", "right", "insideH", "insideV"].map((side) => `<w:${side} w:val="none" w:sz="0" w:space="0" w:color="auto"/>`).join("");

/** XML thân văn bản (một bảng) của khối đầu; người gọi thêm một đoạn trống phía sau nếu cần. */
export function buildHeaderBlock(input: HeaderBlockInput): string {
  const problems = validateHeaderInput(input);
  if (problems.length > 0) throw new Error(problems.join(" "));

  const total = mmToTwips(input.textWidthMm);
  const leftWidth = Math.round(total * LEFT_COLUMN_RATIO);
  const rightWidth = total - leftWidth;
  const mm = (twips: number) => (twips * 25.4) / 1440;

  const parents = input.parentLines.map(clean).filter((line) => line !== "").map((line) => line.toUpperCase());
  const issuer = clean(input.issuer).toUpperCase();
  const issuerLine = Math.min(textWidthMm(issuer, input.sizes.issuer, true), mm(leftWidth)) * input.agencyLineRatio;
  const tieuNguLine = Math.min(textWidthMm(TIEU_NGU, input.sizes.tieuNgu, true), mm(rightWidth));

  const left = [
    ...parents.map((line) => textParagraph(line, input.sizes.parent)),
    textParagraph(issuer, input.sizes.issuer, { bold: true }),
    lineParagraph(issuerLine, leftWidth),
  ];
  const right = [textParagraph(QUOC_HIEU, input.sizes.quocHieu, { bold: true }), textParagraph(TIEU_NGU, input.sizes.tieuNgu, { bold: true }), lineParagraph(tieuNguLine, rightWidth)];

  const numberRow = input.numberRow
    ? `<w:tr>${cell(leftWidth, [textParagraph(input.numberRow.so, input.sizes.so, { beforePt: 6 })])}${cell(rightWidth, [
        textParagraph(input.numberRow.placeDate, input.sizes.placeDate, { italic: true, beforePt: 6 }),
      ])}</w:tr>`
    : "";

  return (
    `<w:tbl><w:tblPr><w:tblW w:w="${total}" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>${NO_BORDERS}</w:tblBorders>` +
    `<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="${leftWidth}"/><w:gridCol w:w="${rightWidth}"/></w:tblGrid>` +
    `<w:tr>${cell(leftWidth, left)}${cell(rightWidth, right)}</w:tr>${numberRow}</w:tbl>`
  );
}

/** Dựng thông số khối đầu từ Tuỳ chọn nhanh: cỡ chữ, tỉ lệ đường kẻ và bề rộng vùng chữ theo lề đang chọn. */
export function headerInputFromOptions(
  profile: StandardProfile,
  values: FormatValues,
  text: { parentLines: readonly string[]; issuer: string; numberRow?: { so: string; placeDate: string } },
): HeaderBlockInput {
  const num = (id: string, fallback: number): number => (typeof values[id] === "number" ? (values[id] as number) : fallback);
  const { widthMm } = profile.page.paper;
  const fixed = profile.pageDefaults.marginsMm;
  return {
    ...text,
    sizes: {
      parent: num("size.tenCoQuanChuQuan", 13),
      issuer: num("size.tenCoQuanBanHanh", 13),
      quocHieu: num("size.quocHieu", 13),
      tieuNgu: num("size.tieuNgu", 14),
      so: num("size.so", 13),
      placeDate: num("size.diaDanhNgayThang", 13),
    },
    agencyLineRatio: num("agencyLineRatio", profile.agencyLineRatio?.max ?? 0.5),
    textWidthMm: widthMm - num("marginLeft", fixed.left) - num("marginRight", fixed.right),
  };
}
