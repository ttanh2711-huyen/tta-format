import type { AssembleInput } from "../ooxml/documentAssembler";
import { BODY_MARKER } from "../ooxml/documentAssembler";
import { buildFooterBlock } from "../ooxml/footerBlock";
import { mmToTwips } from "../ooxml/flatOpc";
import { buildHeaderBlock, escapeXml } from "../ooxml/headerBlock";
import { diemParagraph, multilevelNumberingXml, numberedParagraph, regulationLevels, validateSubItems, type LevelSpec } from "../ooxml/numbering";
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
  /** Số Khoản dựng sẵn trong mỗi Điều (mặc định 1). */
  khoanPerDieu?: number;
  /** Số Điểm dựng sẵn trong mỗi Khoản (mặc định 0; cần có Khoản). */
  diemPerKhoan?: number;
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
  errors.push(...validateSubItems(input.khoanPerDieu ?? 1, input.diemPerKhoan ?? 0));
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

/** Tên style của từng thành phần cấu trúc (mỗi chuẩn một bộ). */
export interface StructureStyles {
  chuong: string;
  tenChuong: string;
  muc: string;
  tenMuc: string;
  dieu: string;
  khoan: string;
  diem: string;
}

export interface StructureSpec {
  chapters: number;
  articlesPerChapter: number;
  withSections: boolean;
  khoanPerDieu?: number;
  diemPerKhoan?: number;
  levels: readonly LevelSpec[];
  styles: StructureStyles;
  /** Dòng đặt trước cấu trúc (đã có ngoặc đơn); trống thì bỏ. */
  attached?: { text: string; sizePt: number };
}

/** Dựng cấu trúc Chương, Mục, Điều, Khoản theo bộ cấp và bộ style cho trước. Tên Chương thứ nhất mang chữ đánh dấu chỗ nội dung. */
export function buildStructureBody(spec: StructureSpec): string {
  const id = (name: string) => styleIdFromName(name);
  const parts: string[] = [];
  if (spec.attached && spec.attached.text !== "") parts.push(attachedParagraph(spec.attached.text, spec.attached.sizePt));
  for (let chapter = 0; chapter < spec.chapters; chapter += 1) {
    parts.push(numberedParagraph(spec.levels, "chuong", id(spec.styles.chuong), ""));
    parts.push(`<w:p><w:pPr><w:pStyle w:val="${id(spec.styles.tenChuong)}"/></w:pPr>${chapter === 0 ? MARKER_RUN : ""}</w:p>`);
    if (spec.withSections) {
      parts.push(numberedParagraph(spec.levels, "muc", id(spec.styles.muc), ""));
      parts.push(`<w:p><w:pPr><w:pStyle w:val="${id(spec.styles.tenMuc)}"/></w:pPr></w:p>`);
    }
    for (let article = 0; article < spec.articlesPerChapter; article += 1) {
      parts.push(numberedParagraph(spec.levels, "dieu", id(spec.styles.dieu), ""));
      for (let k = 0; k < (spec.khoanPerDieu ?? 1); k += 1) {
        parts.push(numberedParagraph(spec.levels, "khoan", id(spec.styles.khoan), ""));
        for (let d = 0; d < (spec.diemPerKhoan ?? 0); d += 1) parts.push(diemParagraph(id(spec.styles.diem), d));
      }
    }
  }
  return parts.join("");
}

/**
 * XML thân dưới tên loại và trích yếu: dòng ban hành kèm theo, các Chương (mỗi Chương: nhãn, tên, tuỳ chọn Mục, các
 * Điều, mỗi Điều một Khoản trống). Tên Chương thứ nhất mang chữ đánh dấu chỗ nội dung nên con trỏ đặt ở đó.
 */
export function buildRegulationBody(input: RegulationInput): string {
  const problems = validateRegulation(input);
  if (problems.length > 0) throw new Error(problems.join(" "));
  const prefix = input.stylePrefix;
  return buildStructureBody({
    chapters: input.chapters,
    articlesPerChapter: input.articlesPerChapter,
    withSections: input.withSections,
    khoanPerDieu: input.khoanPerDieu,
    diemPerKhoan: input.diemPerKhoan,
    levels: regulationLevels(mmToTwips(input.firstLineIndentMm)),
    styles: { chuong: `${prefix}PhanChuong`, tenChuong: `${prefix}TieuDePhanChuong`, muc: `${prefix}MucTieuMuc`, tenMuc: `${prefix}TieuDeMucTieuMuc`, dieu: `${prefix}Dieu`, khoan: `${prefix}Khoan`, diem: `${prefix}Diem` },
    attached: { text: attachedText(input.attachedLine), sizePt: input.sizePt },
  });
}

/** Dựng thông số từ Tuỳ chọn nhanh. */
export function regulationInputFromOptions(profile: StandardProfile, values: FormatValues, text: Pick<RegulationInput, "chapters" | "articlesPerChapter" | "withSections" | "attachedLine" | "khoanPerDieu" | "diemPerKhoan">): RegulationInput {
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
