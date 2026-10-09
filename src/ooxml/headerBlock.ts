import type { FormatValues } from "../settings/formatOptions";
import type { StandardProfile } from "../standards/types";
import { mmToTwips } from "./flatOpc";
import { lineShapeRun } from "./lineShape";
import { textWidthMm } from "./textWidth";

/**
 * Khối đầu văn bản chuẩn Chính quyền (T7): bảng 2 cột không viền, trái là cơ quan chủ quản và cơ quan ban
 * hành, phải là quốc hiệu và tiêu ngữ; mỗi bên có đường kẻ dưới bằng viền của một đoạn riêng (không dùng hình
 * nổi nên đi theo văn bản). Hàng số/ký hiệu, địa danh và V/v thêm ở T8, T9.
 */

export const QUOC_HIEU = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
export const TIEU_NGU = "Độc lập - Tự do - Hạnh phúc";
export const MAX_PARENT_LINES = 3;

/**
 * Hàng "Kính gửi" trong bảng đầu văn bản: cả hai ô bắt đầu từ dòng thứ hai (dòng đầu để trống). Một nơi nhận nằm ngang
 * hàng với nhãn; từ hai nơi nhận trở lên, danh sách nằm ngay dưới dòng của nhãn.
 */
export interface RecipientsRow {
  label: string;
  italicLabel?: boolean;
  /** Mỗi phần tử một dòng nơi nhận, đã có dấu gạch đầu dòng nếu cần. */
  lines: readonly string[];
  sizePt: number;
}

const DASH_HANGING_MM = 4;

export function recipientsRowXml(row: RecipientsRow, leftWidth: number, rightWidth: number): string {
  const blank = () => textParagraph("", row.sizePt);
  const label = textParagraph(row.label, row.sizePt, { italic: row.italicLabel, align: "right", afterPt: 12 });
  const names = row.lines.map((line, index) =>
    textParagraph(line, row.sizePt, { align: "left", hangingMm: line.startsWith("- ") ? DASH_HANGING_MM : undefined, afterPt: index === row.lines.length - 1 ? 12 : 0 }),
  );
  // Một nơi nhận đứng cùng dòng với nhãn; từ hai nơi nhận trở lên thì danh sách bắt đầu ở dòng ngay dưới nhãn.
  const gap = row.lines.length >= 2 ? [blank()] : [];
  return `<w:tr>${cell(leftWidth, [blank(), label])}${cell(rightWidth, [blank(), ...gap, ...names])}</w:tr>`;
}

/** Tỉ lệ bề rộng cột trái của bảng đầu văn bản (theo file mẫu). */
export const LEFT_COLUMN_RATIO = 0.45;
const FONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>';

export interface HeaderBlockInput {
  /** Cơ quan chủ quản, từ 0 đến 3 dòng; dòng trống bị bỏ. */
  parentLines: readonly string[];
  issuer: string;
  sizes: { parent: number; issuer: number; quocHieu: number; tieuNgu: number; so: number; placeDate: number; subject?: number };
  /** Hàng thứ hai: dòng số, ký hiệu (trái) và địa danh, ngày tháng (phải), đã dựng sẵn thành chữ (T8). */
  numberRow?: { so: string; placeDate: string };
  /** Hàng thứ ba của công văn: dòng "V/v ..." ở ô trái, ô phải để trống (T9). */
  subjectRow?: { text: string };
  /** Hàng "Kính gửi" của công văn: nhãn ở ô trái (canh phải), các nơi nhận ở ô phải (canh trái). */
  recipientsRow?: RecipientsRow;
  /** Độ dài đường kẻ dưới tên cơ quan so với độ dài dòng chữ. */
  agencyLineRatio: number;
  /** Bề rộng vùng chữ (mm) mà bảng phải vừa. */
  textWidthMm: number;
}

function clean(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

export function escapeXml(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function validateHeaderInput(input: HeaderBlockInput): string[] {
  const errors: string[] = [];
  if (clean(input.issuer) === "") errors.push("Hãy nhập tên cơ quan ban hành.");
  if (input.parentLines.length > MAX_PARENT_LINES) errors.push(`Cơ quan chủ quản tối đa ${MAX_PARENT_LINES} dòng.`);
  return errors;
}

export interface TextStyle {
  bold?: boolean;
  italic?: boolean;
  beforePt?: number;
  afterPt?: number;
  /** Căn lề của đoạn; mặc định căn giữa. */
  align?: "center" | "left" | "right";
  /** Thụt treo (mm) cho dòng gạch đầu dòng, để các dòng xuống hàng thẳng hàng sau dấu gạch. */
  hangingMm?: number;
  shape?: string;
}

export function textParagraph(text: string, sizePt: number, style: TextStyle = {}): string {
  const half = sizePt * 2;
  const before = (style.beforePt ?? 0) * 20;
  const after = (style.afterPt ?? 0) * 20;
  const indent = style.hangingMm ? `<w:ind w:left="${mmToTwips(style.hangingMm)}" w:hanging="${mmToTwips(style.hangingMm)}"/>` : "";
  return (
    `<w:p><w:pPr><w:spacing w:before="${before}" w:after="${after}" w:line="240" w:lineRule="auto"/>${indent}<w:jc w:val="${style.align ?? "center"}"/></w:pPr>` +
    `${style.shape ?? ""}<w:r><w:rPr>${FONT}${style.bold ? "<w:b/>" : ""}${style.italic ? "<w:i/>" : ""}<w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`
  );
}

/** Đoạn rỗng cao 6 pt có viền dưới, thụt hai bên để đường kẻ dài `lineMm` và nằm giữa ô rộng `cellTwips`. */
export function lineParagraph(lineMm: number, cellTwips: number): string {
  const inset = Math.max(0, Math.round((cellTwips - mmToTwips(lineMm)) / 2));
  return (
    `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>` +
    `<w:spacing w:before="0" w:after="0" w:line="120" w:lineRule="exact"/>` +
    `<w:ind w:left="${inset}" w:right="${inset}"/><w:jc w:val="center"/></w:pPr></w:p>`
  );
}

export function cell(widthTwips: number, paragraphs: string[]): string {
  return `<w:tc><w:tcPr><w:tcW w:w="${widthTwips}" w:type="dxa"/></w:tcPr>${paragraphs.join("")}</w:tc>`;
}

const NO_BORDERS = ["top", "left", "bottom", "right", "insideH", "insideV"].map((side) => `<w:${side} w:val="none" w:sz="0" w:space="0" w:color="auto"/>`).join("");

/** Bảng 2 cột không viền, lề ô 0, bố cục cố định, rộng `total` twips; `rowsXml` là các hàng. */
export function borderlessTable(total: number, leftWidth: number, rowsXml: string): string {
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="${total}" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>${NO_BORDERS}</w:tblBorders>` +
    `<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="${leftWidth}"/><w:gridCol w:w="${total - leftWidth}"/></w:tblGrid>${rowsXml}</w:tbl>`
  );
}

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

  const issuerShape = lineShapeRun({ lengthMm: issuerLine, columnTwips: leftWidth, textSizePt: input.sizes.issuer, id: 101 });
  const tieuNguShape = lineShapeRun({ lengthMm: tieuNguLine, columnTwips: rightWidth, textSizePt: input.sizes.tieuNgu, id: 102 });
  const left = [...parents.map((line) => textParagraph(line, input.sizes.parent)), textParagraph(issuer, input.sizes.issuer, { bold: true, shape: issuerShape })];
  const right = [textParagraph(QUOC_HIEU, input.sizes.quocHieu, { bold: true }), textParagraph(TIEU_NGU, input.sizes.tieuNgu, { bold: true, shape: tieuNguShape })];

  const numberRow = input.numberRow
    ? `<w:tr>${cell(leftWidth, [textParagraph(input.numberRow.so, input.sizes.so, { beforePt: 6 })])}${cell(rightWidth, [
        textParagraph(input.numberRow.placeDate, input.sizes.placeDate, { italic: true, beforePt: 6 }),
      ])}</w:tr>`
    : "";

  const subjectRow = input.subjectRow
    ? `<w:tr>${cell(leftWidth, [textParagraph(input.subjectRow.text, input.sizes.subject ?? 13, { beforePt: 6 })])}${cell(rightWidth, [textParagraph("", 12)])}</w:tr>`
    : "";

  return borderlessTable(total, leftWidth, `<w:tr>${cell(leftWidth, left)}${cell(rightWidth, right)}</w:tr>${numberRow}${subjectRow}${input.recipientsRow ? recipientsRowXml(input.recipientsRow, leftWidth, rightWidth) : ""}`);
}


/** Dựng thông số khối đầu từ Tuỳ chọn nhanh: cỡ chữ, tỉ lệ đường kẻ và bề rộng vùng chữ theo lề đang chọn. */
export function headerInputFromOptions(
  profile: StandardProfile,
  values: FormatValues,
  text: { parentLines: readonly string[]; issuer: string; numberRow?: { so: string; placeDate: string }; subjectRow?: { text: string }; recipientsRow?: RecipientsRow },
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
      subject: num("size.trichYeuCongVan", 13),
    },
    agencyLineRatio: num("agencyLineRatio", profile.agencyLineRatio?.max ?? 0.5),
    textWidthMm: widthMm - num("marginLeft", fixed.left) - num("marginRight", fixed.right),
  };
}
