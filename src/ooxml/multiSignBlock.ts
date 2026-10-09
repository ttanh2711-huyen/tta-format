import type { FormatValues } from "../settings/formatOptions";
import type { StandardProfile } from "../standards/types";
import { mmToTwips } from "./flatOpc";
import { escapeXml, textParagraph } from "./headerBlock";
import { recipientLines } from "./footerBlock";

/**
 * Khối ký nhiều bên (T19): bảng không viền, các cột bằng nhau (2 đến 4 bên), mỗi cột gồm tên bên hoặc vai trò (in hoa
 * đậm), chức vụ (in hoa đậm), chỗ ký và họ tên (đậm). Tên dài tự xuống dòng trong cột của mình nên bố cục không vỡ.
 * Hàng thứ hai (tuỳ chọn) là "Nơi nhận" nằm sát lề trái, trải hết các cột; loại văn bản không yêu cầu thì bỏ.
 */

export const MIN_PARTIES = 2;
export const MAX_PARTIES = 4;
const MAX_SIGN_GAP = 12;
export const MAX_POSITION_LINES = 3;
const FONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>';
const NO_BORDERS = ["top", "left", "bottom", "right", "insideH", "insideV"].map((side) => `<w:${side} w:val="none" w:sz="0" w:space="0" w:color="auto"/>`).join("");

export interface SignParty {
  /** Tên bên hoặc vai trò ký, ví dụ "ĐẠI DIỆN BÊN A", "CHỦ TỌA", "THƯ KÝ". */
  title: string;
  /** Chức vụ, 0 đến 3 dòng. */
  positionLines: readonly string[];
  /** Họ tên; để trống nếu chưa biết (chừa chỗ). */
  signerName: string;
}

export interface MultiSignInput {
  parties: readonly SignParty[];
  signGapLines: number;
  sizes: { title: number; name: number; noiNhanTieuDe: number; noiNhanDanhSach: number };
  textWidthMm: number;
  /** Nơi nhận; rỗng thì không có hàng "Nơi nhận". */
  recipients: readonly string[];
  luu: string;
}

function clean(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

export function validateMultiSign(input: MultiSignInput): string[] {
  const errors: string[] = [];
  if (input.parties.length < MIN_PARTIES || input.parties.length > MAX_PARTIES) errors.push(`Khối ký có từ ${MIN_PARTIES} đến ${MAX_PARTIES} bên.`);
  input.parties.forEach((party, index) => {
    if (clean(party.title) === "") errors.push(`Bên thứ ${index + 1}: hãy nhập tên bên hoặc vai trò ký.`);
    if (party.positionLines.length > MAX_POSITION_LINES) errors.push(`Bên thứ ${index + 1}: chức vụ tối đa ${MAX_POSITION_LINES} dòng.`);
  });
  if (!Number.isInteger(input.signGapLines) || input.signGapLines < 1 || input.signGapLines > MAX_SIGN_GAP) errors.push(`Số dòng trống chỗ ký phải từ 1 đến ${MAX_SIGN_GAP}.`);
  return errors;
}

function leftParagraph(text: string, sizePt: number, style: { bold?: boolean; italic?: boolean }): string {
  const half = sizePt * 2;
  const props = `${FONT}${style.bold ? "<w:b/>" : ""}${style.italic ? "<w:i/>" : ""}<w:sz w:val="${half}"/><w:szCs w:val="${half}"/>`;
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/><w:rPr>${props}</w:rPr></w:pPr><w:r><w:rPr>${props}</w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function cell(width: number, paragraphs: string[], span = 1): string {
  const grid = span > 1 ? `<w:gridSpan w:val="${span}"/>` : "";
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${grid}</w:tcPr>${paragraphs.join("")}</w:tc>`;
}

/** Dựng thông số từ Tuỳ chọn nhanh: cỡ chữ, số dòng trống chỗ ký, bề rộng vùng chữ. */
export function multiSignFromOptions(profile: StandardProfile, values: FormatValues, text: Pick<MultiSignInput, "parties" | "recipients" | "luu">): MultiSignInput {
  const num = (id: string, fallback: number): number => (typeof values[id] === "number" ? (values[id] as number) : fallback);
  const fixed = profile.pageDefaults.marginsMm;
  return {
    ...text,
    signGapLines: num("signGap", 6),
    sizes: { title: num("size.quyenHanChucVu", 14), name: num("size.hoTenNguoiKy", 14), noiNhanTieuDe: num("size.noiNhanTieuDe", 12), noiNhanDanhSach: num("size.noiNhanDanhSach", 11) },
    textWidthMm: profile.page.paper.widthMm - num("marginLeft", fixed.left) - num("marginRight", fixed.right),
  };
}

/** Đọc các dòng "Tên bên | Chức vụ | Họ tên" thành danh sách bên ký; dòng trống bị bỏ. */
export function parseSignLines(text: string): SignParty[] {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line !== "")
    .map((line) => {
      const [title = "", position = "", name = ""] = line.split("|").map((part) => part.trim());
      return { title, positionLines: position === "" ? [] : [position], signerName: name };
    });
}

/** XML thân văn bản của khối ký nhiều bên và một đoạn đệm cao 1 pt phía sau (như khối cuối thường). */
export function buildMultiSignBlock(input: MultiSignInput): string {
  const problems = validateMultiSign(input);
  if (problems.length > 0) throw new Error(problems.join(" "));
  const total = mmToTwips(input.textWidthMm);
  const count = input.parties.length;
  const base = Math.floor(total / count);
  const widths = input.parties.map((_, index) => (index === count - 1 ? total - base * (count - 1) : base));
  const { sizes } = input;

  const columns = input.parties.map((party, index) => {
    const lines = [clean(party.title).toUpperCase(), ...party.positionLines.map((line) => clean(line).toUpperCase()).filter((line) => line !== "")];
    const paragraphs = [
      ...lines.map((line) => textParagraph(line, sizes.title, { bold: true })),
      ...Array.from({ length: input.signGapLines }, () => textParagraph("", sizes.title, { bold: true })),
      textParagraph(clean(party.signerName), sizes.name, { bold: true }),
    ];
    return cell(widths[index], paragraphs);
  });

  const lines = recipientLines(input.recipients, input.luu);
  const recipientsRow =
    lines.length === 0
      ? ""
      : `<w:tr>${cell(total, [leftParagraph("Nơi nhận:", sizes.noiNhanTieuDe, { bold: true, italic: true }), ...lines.map((line) => leftParagraph(line, sizes.noiNhanDanhSach, {}))], count)}</w:tr>`;

  const tail = `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr>${FONT}<w:sz w:val="2"/></w:rPr></w:pPr></w:p>`;
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="${total}" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>${NO_BORDERS}</w:tblBorders>` +
    `<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
    `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>` +
    `<w:tr>${columns.join("")}</w:tr>${recipientsRow}</w:tbl>${tail}`
  );
}
