import type { AssembleInput } from "../ooxml/documentAssembler";
import { BODY_MARKER } from "../ooxml/documentAssembler";
import { buildFooterBlock } from "../ooxml/footerBlock";
import { mmToTwips } from "../ooxml/flatOpc";
import { buildHeaderBlock, escapeXml } from "../ooxml/headerBlock";
import { multilevelNumberingXml, numberedParagraph, regulationLevels } from "../ooxml/numbering";
import { styleIdFromName } from "../ooxml/styleXml";
import { buildTitleBody } from "../ooxml/titleBlock";
import type { FormatValues } from "../settings/formatOptions";
import type { StandardProfile } from "../standards/types";

/**
 * Quy chế, quy định (T18): tên loại và trích yếu, dòng "(Ban hành kèm theo ...)", rồi cấu trúc Chương, Mục, Điều,
 * Khoản. Nhãn "Chương I", "Mục 1", "Điều 1." do danh sách nhiều cấp tự đánh số (số Điều chạy liên tục qua các
 * Chương); tên Chương và tên Mục là đoạn riêng ngay dưới nhãn (in hoa, đậm); tên Điều nằm cùng dòng với nhãn.
 */

export const MAX_CHAPTERS = 20;
export const MAX_ARTICLES_PER_CHAPTER = 20;

export interface RegulationInput {
  chapters: number;
  /** Số Điều dựng sẵn trong mỗi Chương (người dùng thêm bớt sau). */
  articlesPerChapter: number;
  /** Mỗi Chương có một Mục trước các Điều. */
  withSections: boolean;
  /** Dòng "(Ban hành kèm theo Quyết định số ... ngày ... của ...)"; trống thì bỏ. */
  attachedLine: string;
  stylePrefix: string;
  firstLineIndentMm: number;
  sizePt: number;
}

function clean(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

export function validateRegulation(input: RegulationInput): string[] {
  const errors: string[] = [];
  if (!Number.isInteger(input.chapters) || input.chapters < 1 || input.chapters > MAX_CHAPTERS) errors.push(`Số Chương phải từ 1 đến ${MAX_CHAPTERS}.`);
  if (!Number.isInteger(input.articlesPerChapter) || input.articlesPerChapter < 1 || input.articlesPerChapter > MAX_ARTICLES_PER_CHAPTER) {
    errors.push(`Số Điều mỗi Chương phải từ 1 đến ${MAX_ARTICLES_PER_CHAPTER}.`);
  }
  return errors;
}

/** Dòng ban hành kèm theo, đặt trong ngoặc đơn nếu người dùng chưa gõ. */
export function attachedText(text: string): string {
  const value = clean(text);
  if (value === "") return "";
  return value.startsWith("(") ? value : `(${value})`;
}

const MARKER_RUN = `<w:r><w:t>${BODY_MARKER}</w:t></w:r>`;

export function regulationNumberingXml(input: Pick<RegulationInput, "firstLineIndentMm">): string {
  return multilevelNumberingXml(regulationLevels(mmToTwips(input.firstLineIndentMm)));
}

function attachedParagraph(text: string, sizePt: number): string {
  const half = sizePt * 2;
  return `<w:p><w:pPr><w:keepNext/><w:spacing w:before="0" w:after="120"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:i/><w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

/**
 * XML thân dưới tên loại và trích yếu: dòng ban hành kèm theo, các Chương (mỗi Chương: nhãn, tên, tuỳ chọn Mục, các
 * Điều, mỗi Điều một Khoản trống). Tên Chương thứ nhất mang chữ đánh dấu chỗ nội dung nên con trỏ đặt ở đó.
 */
export function buildRegulationBody(input: RegulationInput): string {
  const problems = validateRegulation(input);
  if (problems.length > 0) throw new Error(problems.join(" "));
  const levels = regulationLevels(mmToTwips(input.firstLineIndentMm));
  const id = (element: string) => styleIdFromName(`${input.stylePrefix}${element}`);
  const parts: string[] = [];
  const attached = attachedText(input.attachedLine);
  if (attached !== "") parts.push(attachedParagraph(attached, input.sizePt));
  for (let chapter = 0; chapter < input.chapters; chapter += 1) {
    parts.push(numberedParagraph(levels, "chuong", id("PhanChuong"), ""));
    parts.push(`<w:p><w:pPr><w:pStyle w:val="${id("TieuDePhanChuong")}"/></w:pPr>${chapter === 0 ? MARKER_RUN : ""}</w:p>`);
    if (input.withSections) {
      parts.push(numberedParagraph(levels, "muc", id("MucTieuMuc"), ""));
      parts.push(`<w:p><w:pPr><w:pStyle w:val="${id("TieuDeMucTieuMuc")}"/></w:pPr></w:p>`);
    }
    for (let article = 0; article < input.articlesPerChapter; article += 1) {
      parts.push(numberedParagraph(levels, "dieu", id("Dieu"), ""));
      parts.push(numberedParagraph(levels, "khoan", id("Khoan"), ""));
    }
  }
  return parts.join("");
}

/** Dựng thông số từ Tuỳ chọn nhanh. */
export function regulationInputFromOptions(profile: StandardProfile, values: FormatValues, text: Pick<RegulationInput, "chapters" | "articlesPerChapter" | "withSections" | "attachedLine">): RegulationInput {
  const num = (id: string, fallback: number): number => (typeof values[id] === "number" ? (values[id] as number) : fallback);
  return {
    ...text,
    stylePrefix: profile.styles.prefix,
    firstLineIndentMm: num("firstLineIndent", profile.styles.defaults.firstLineIndentMm),
    sizePt: num("fontSize", profile.styles.defaults.fontSizePt),
  };
}

/** Thân văn bản đầy đủ của quy chế, quy định: khối đầu, tên loại và trích yếu, cấu trúc, khối cuối. */
export function assembleRegulationBody(input: AssembleInput, regulation: RegulationInput): string {
  if (!input.title || input.title.variant !== "named") throw new Error("Quy chế, quy định phải dùng khối tên loại và trích yếu.");
  return buildHeaderBlock(input.header) + buildTitleBody(input.title) + buildRegulationBody(regulation) + buildFooterBlock(input.footer);
}
