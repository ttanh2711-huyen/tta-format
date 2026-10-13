import { mmToTwips } from "../../../ooxml/flatOpc";
import { multilevelNumberingXml, validateSubItems } from "../../../ooxml/numbering";
import type { FormatValues } from "../../../settings/formatOptions";
import { buildStructureBody } from "../../../templates/quyChe";
import type { StandardProfile } from "../../types";
import { buildPartyFooterBlock } from "../footerBlock";
import { buildPartyHeaderBlock } from "../headerBlock";
import { partyRegulationLevels, type PartyNumbering, type PartyTopLevel } from "../numbering";
import type { PartyAssembleInput } from "../assembleDocument";
import { buildPartyTitleBody } from "../titleBlock";

/**
 * Nghị quyết, quyết định, quy định, quy chế của Đảng (TD6): khối đầu, thể loại, trích yếu, 5 dấu gạch nối, căn cứ (từ
 * TD3), rồi cấu trúc Phần hoặc Chương, Mục, Điều, Khoản tự đánh số theo bảng mục 6 của Phụ lục 2 của HD 05.
 */

export const MAX_CHAPTERS = 20;
export const MAX_ARTICLES_PER_CHAPTER = 20;

export interface PartyRegulationInput {
  chapters: number;
  articlesPerChapter: number;
  withSections: boolean;
  /** Số Khoản dựng sẵn trong mỗi Điều (mặc định 1). */
  khoanPerDieu?: number;
  /** Số Điểm dựng sẵn trong mỗi Khoản (mặc định 0; cần có Khoản). */
  diemPerKhoan?: number;
  top: PartyTopLevel;
  numbering: PartyNumbering;
  stylePrefix: string;
  firstLineIndentMm: number;
}

export function validatePartyRegulation(input: PartyRegulationInput): string[] {
  const errors: string[] = [];
  if (!Number.isInteger(input.chapters) || input.chapters < 1 || input.chapters > MAX_CHAPTERS) errors.push(`Số ${input.top} phải từ 1 đến ${MAX_CHAPTERS}.`);
  if (!Number.isInteger(input.articlesPerChapter) || input.articlesPerChapter < 1 || input.articlesPerChapter > MAX_ARTICLES_PER_CHAPTER) {
    errors.push(`Số Điều mỗi ${input.top} phải từ 1 đến ${MAX_ARTICLES_PER_CHAPTER}.`);
  }
  errors.push(...validateSubItems(input.khoanPerDieu ?? 1, input.diemPerKhoan ?? 0));
  return errors;
}

function levelsOf(input: PartyRegulationInput) {
  return partyRegulationLevels(input.top, input.numbering, mmToTwips(input.firstLineIndentMm));
}

export function partyRegulationNumberingXml(input: PartyRegulationInput): string {
  return multilevelNumberingXml(levelsOf(input));
}

/** XML cấu trúc dưới thể loại, trích yếu và căn cứ; tên Phần hoặc Chương thứ nhất mang chữ đánh dấu chỗ nội dung. */
export function buildPartyRegulationBody(input: PartyRegulationInput): string {
  const problems = validatePartyRegulation(input);
  if (problems.length > 0) throw new Error(problems.join(" "));
  const p = input.stylePrefix;
  return buildStructureBody({
    chapters: input.chapters,
    articlesPerChapter: input.articlesPerChapter,
    withSections: input.withSections,
    khoanPerDieu: input.khoanPerDieu,
    diemPerKhoan: input.diemPerKhoan,
    levels: levelsOf(input),
    styles: { chuong: `${p}PhanChuong`, tenChuong: `${p}TenPhanChuong`, muc: `${p}MucSo`, tenMuc: `${p}TenMuc`, dieu: `${p}Dieu`, khoan: `${p}Khoan`, diem: `${p}Diem` },
  });
}

/** Dựng thông số từ Tuỳ chọn nhanh. */
export function partyRegulationFromOptions(
  profile: StandardProfile,
  values: FormatValues,
  text: Pick<PartyRegulationInput, "chapters" | "articlesPerChapter" | "withSections" | "top" | "numbering" | "khoanPerDieu" | "diemPerKhoan">,
): PartyRegulationInput {
  const indent = values.firstLineIndent;
  return { ...text, stylePrefix: profile.styles.prefix, firstLineIndentMm: typeof indent === "number" ? indent : profile.styles.defaults.firstLineIndentMm };
}

/** Thân văn bản Đảng đầy đủ: khối đầu, thể loại và căn cứ, cấu trúc, khối cuối. */
export function assemblePartyRegulationBody(input: PartyAssembleInput, regulation: PartyRegulationInput): string {
  if (!input.title || input.title.variant === "congVan") throw new Error("Văn bản có cấu trúc Phần, Chương, Điều phải có thể loại và trích yếu.");
  return buildPartyHeaderBlock(input.header) + buildPartyTitleBody(input.title) + buildPartyRegulationBody(regulation) + buildPartyFooterBlock(input.footer);
}
