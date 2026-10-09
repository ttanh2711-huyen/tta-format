import type { AssembleInput } from "../ooxml/documentAssembler";
import { BODY_MARKER } from "../ooxml/documentAssembler";
import { mmToTwips } from "../ooxml/flatOpc";
import { buildHeaderBlock, escapeXml } from "../ooxml/headerBlock";
import { buildMultiSignBlock, MAX_PARTIES, MIN_PARTIES, type MultiSignInput } from "../ooxml/multiSignBlock";
import { legalNumberingXml, legalParagraph } from "../ooxml/numbering";
import { styleIdFromName } from "../ooxml/styleXml";
import { buildTitleBody } from "../ooxml/titleBlock";

/**
 * Hợp đồng, bản ghi nhớ, bản thỏa thuận (T19). Phụ lục III không có mẫu cho ba loại này nên đây là bố cục đề xuất
 * (docs/nd30-ambiguities.md, mục 4): khối đầu và tên loại như Mẫu 1.4, dòng mở đầu, danh sách các bên, các Điều đánh
 * số tự động, khối ký nhiều bên, không có "Nơi nhận" cố định.
 */

const LETTERS = ["A", "B", "C", "D"];

export interface AgreementInput {
  /** Số bên, 2 đến 4. */
  parties: number;
  articleCount: number;
  stylePrefix: string;
  firstLineIndentMm: number;
}

export const MAX_ARTICLES = 30;

export function validateAgreement(input: AgreementInput): string[] {
  const errors: string[] = [];
  if (!Number.isInteger(input.parties) || input.parties < MIN_PARTIES || input.parties > MAX_PARTIES) errors.push(`Số bên phải từ ${MIN_PARTIES} đến ${MAX_PARTIES}.`);
  if (!Number.isInteger(input.articleCount) || input.articleCount < 1 || input.articleCount > MAX_ARTICLES) errors.push(`Số Điều phải từ 1 đến ${MAX_ARTICLES}.`);
  return errors;
}

function plain(text: string, styleId: string, boldLabel?: string, marker = false): string {
  const label = boldLabel ? `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${escapeXml(boldLabel)}</w:t></w:r>` : "";
  const body = text === "" ? "" : `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
  const mark = marker ? `<w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>${BODY_MARKER}</w:t></w:r>` : "";
  return `<w:p><w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>${label}${body}${mark}</w:p>`;
}

/** Nhãn ký mặc định của từng bên: "Đại diện bên A", "Đại diện bên B"... */
export function defaultAgreementParties(count: number): MultiSignInput["parties"] {
  return LETTERS.slice(0, count).map((letter) => ({ title: `Đại diện bên ${letter}`, positionLines: [], signerName: "" }));
}

export function agreementNumberingXml(input: Pick<AgreementInput, "firstLineIndentMm">): string {
  return legalNumberingXml(mmToTwips(input.firstLineIndentMm));
}

/** XML dưới tên loại: dòng mở đầu, các bên, câu dẫn vào các Điều, các Điều. Chữ đánh dấu chỗ nội dung nằm sau nhãn của bên A. */
export function buildAgreementBody(input: AgreementInput): string {
  const problems = validateAgreement(input);
  if (problems.length > 0) throw new Error(problems.join(" "));
  const style = styleIdFromName(`${input.stylePrefix}NoiDung`);
  const parts = [plain("Hôm nay, ngày … tháng … năm …, tại …, chúng tôi gồm:", style)];
  LETTERS.slice(0, input.parties).forEach((letter, index) => parts.push(plain("", style, `Bên ${letter}:`, index === 0)));
  parts.push(plain("Các bên thống nhất thỏa thuận như sau:", style));
  for (let i = 0; i < input.articleCount; i += 1) parts.push(legalParagraph("dieu", style, ""));
  return parts.join("");
}

/** Thân văn bản đầy đủ: khối đầu, tên loại, nội dung, khối ký nhiều bên. */
export function assembleAgreementBody(input: Pick<AssembleInput, "header" | "title">, agreement: AgreementInput, sign: MultiSignInput): string {
  if (!input.title || input.title.variant !== "named") throw new Error("Hợp đồng, bản ghi nhớ, bản thỏa thuận phải dùng khối tên loại và trích yếu.");
  return buildHeaderBlock(input.header) + buildTitleBody(input.title) + buildAgreementBody(agreement) + buildMultiSignBlock(sign);
}
