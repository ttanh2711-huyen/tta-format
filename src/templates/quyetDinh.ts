import type { AssembleInput } from "../ooxml/documentAssembler";
import { BODY_MARKER } from "../ooxml/documentAssembler";
import { buildFooterBlock } from "../ooxml/footerBlock";
import { mmToTwips } from "../ooxml/flatOpc";
import { buildHeaderBlock, escapeXml } from "../ooxml/headerBlock";
import { buildTitleBody } from "../ooxml/titleBlock";
import { diemParagraph, legalNumberingXml, legalParagraph, validateSubItems } from "../ooxml/numbering";
import { styleIdFromName } from "../ooxml/styleXml";
import type { FormatValues } from "../settings/formatOptions";
import type { StandardProfile } from "../standards/types";

/**
 * Nghị quyết và quyết định cá biệt (T16): phần dưới tên loại và trích yếu gồm thẩm quyền ban hành, các căn cứ
 * (nghiêng), dòng "QUYẾT ĐỊNH:" (hoặc "QUYẾT NGHỊ:") và các Điều đánh số tự động. Mẫu 1.1 nghị quyết cá biệt,
 * Mẫu 1.2 quyết định quy định trực tiếp, Mẫu 1.3 quyết định quy định gián tiếp (ban hành kèm theo văn bản khác).
 */

export type DecisionKind = "nghiQuyet" | "quyetDinhTrucTiep" | "quyetDinhGianTiep";

export interface DecisionKindInfo {
  kind: DecisionKind;
  label: string;
  /** Tên loại ghi dưới khối đầu (in hoa). */
  typeName: string;
  /** Dòng chốt trước các Điều. */
  verdict: string;
  template: string;
}

export const DECISION_KINDS: readonly DecisionKindInfo[] = [
  { kind: "nghiQuyet", label: "Nghị quyết cá biệt", typeName: "Nghị quyết", verdict: "QUYẾT NGHỊ:", template: "Mẫu 1.1" },
  { kind: "quyetDinhTrucTiep", label: "Quyết định cá biệt (quy định trực tiếp)", typeName: "Quyết định", verdict: "QUYẾT ĐỊNH:", template: "Mẫu 1.2" },
  { kind: "quyetDinhGianTiep", label: "Quyết định cá biệt (quy định gián tiếp)", typeName: "Quyết định", verdict: "QUYẾT ĐỊNH:", template: "Mẫu 1.3" },
];

export function decisionKindInfo(kind: DecisionKind): DecisionKindInfo {
  const info = DECISION_KINDS.find((item) => item.kind === kind);
  if (!info) throw new Error(`Không có loại "${kind}".`);
  return info;
}

export const MAX_ARTICLES = 30;
/** Câu mở Điều 1 của quyết định quy định gián tiếp; con trỏ nằm ngay sau câu này. */
export const INDIRECT_LEAD = "Ban hành kèm theo Quyết định này ";

export interface DecisionInput {
  kind: DecisionKind;
  /** Thẩm quyền ban hành (chức danh hoặc tên cơ quan), ví dụ "BỘ TRƯỞNG BỘ NỘI VỤ"; in hoa đậm, canh giữa. */
  authority: string;
  /** Các căn cứ, mỗi căn cứ một dòng, không cần dấu cuối dòng. */
  basis: readonly string[];
  articleCount: number;
  /** Số Khoản dựng sẵn trong mỗi Điều (mặc định 0). */
  khoanPerDieu?: number;
  /** Số Điểm dựng sẵn trong mỗi Khoản (mặc định 0; cần có Khoản). */
  diemPerKhoan?: number;
  /** Tiền tố style của chuẩn, ví dụ "ND30_". */
  stylePrefix: string;
  sizes: { authority: number; basis: number; verdict: number };
  /** Thụt dòng đầu của nội dung (mm); nhãn "Điều 1." nằm trong dòng đầu. */
  firstLineIndentMm: number;
}

function clean(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

export function validateDecision(input: DecisionInput): string[] {
  const errors: string[] = [];
  if (clean(input.authority) === "") errors.push("Hãy nhập thẩm quyền ban hành (chức danh hoặc tên cơ quan).");
  if (input.basis.map(clean).filter((line) => line !== "").length === 0) errors.push("Hãy nhập ít nhất một căn cứ.");
  if (!Number.isInteger(input.articleCount) || input.articleCount < 1 || input.articleCount > MAX_ARTICLES) errors.push(`Số Điều phải từ 1 đến ${MAX_ARTICLES}.`);
  errors.push(...validateSubItems(input.khoanPerDieu ?? 0, input.diemPerKhoan ?? 0));
  return errors;
}

/**
 * Căn cứ hoàn chỉnh: mỗi căn cứ một dòng, cuối dòng ";" và căn cứ cuối kết thúc bằng ",", vì sau đó là dòng
 * "QUYẾT ĐỊNH:". (Cách ghi dấu căn cứ cuối theo thông lệ; cần đối chiếu Phụ lục I, mục II.6.a.)
 */
export function basisLines(basis: readonly string[]): string[] {
  const entries = basis.map((line) => clean(line).replace(/[;.,]+$/, "")).filter((line) => line !== "");
  // Phụ lục I, II.6: cuối mỗi căn cứ có dấu chấm phẩy, dòng cuối cùng kết thúc bằng dấu chấm.
  return entries.map((entry, index) => `${entry}${index === entries.length - 1 ? "." : ";"}`);
}

function plain(text: string, sizePt: number, options: { bold?: boolean; italic?: boolean; align: "center" | "both"; styleId?: string; keepNext?: boolean; beforePt?: number; afterPt?: number }): string {
  const half = sizePt * 2;
  const style = options.styleId ? `<w:pStyle w:val="${options.styleId}"/>` : "";
  const keep = options.keepNext ? "<w:keepNext/>" : "";
  const spacing = options.beforePt === undefined && options.afterPt === undefined ? "" : `<w:spacing w:before="${(options.beforePt ?? 0) * 20}" w:after="${(options.afterPt ?? 0) * 20}"/>`;
  const rpr = `${options.bold ? "<w:b/>" : ""}${options.italic ? "<w:i/>" : ""}<w:sz w:val="${half}"/><w:szCs w:val="${half}"/>`;
  return `<w:p><w:pPr>${style}${keep}${spacing}<w:jc w:val="${options.align}"/></w:pPr><w:r><w:rPr>${rpr}</w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

const MARKER_RUN = `<w:r><w:t>${BODY_MARKER}</w:t></w:r>`;

/** Phần numbering.xml đi kèm thân văn bản của quyết định (chèn qua `flatOpc({ numberingXml })`). */
export function decisionNumberingXml(input: Pick<DecisionInput, "firstLineIndentMm">): string {
  return legalNumberingXml(mmToTwips(input.firstLineIndentMm));
}

/**
 * XML thân văn bản dưới tên loại và trích yếu: thẩm quyền ban hành, căn cứ, dòng chốt, các Điều. Điều 1 mang chữ
 * đánh dấu chỗ nội dung (con trỏ sẽ đặt ở đó). Điều dùng style nội dung với nhãn "Điều N." in đậm do danh sách đánh
 * số; nội dung điều không đậm.
 */
export function buildDecisionBody(input: DecisionInput): string {
  const problems = validateDecision(input);
  if (problems.length > 0) throw new Error(problems.join(" "));
  const info = decisionKindInfo(input.kind);
  const body = styleIdFromName(`${input.stylePrefix}NoiDung`);
  const basisStyle = styleIdFromName(`${input.stylePrefix}CanCu`);
  const parts = [plain(clean(input.authority).toUpperCase(), input.sizes.authority, { bold: true, align: "center", keepNext: true, beforePt: 6, afterPt: 6 })];
  for (const line of basisLines(input.basis)) parts.push(plain(line, input.sizes.basis, { italic: true, align: "both", styleId: basisStyle }));
  parts.push(plain(info.verdict, input.sizes.verdict, { bold: true, align: "center", keepNext: true, beforePt: 6, afterPt: 6 }));
  const indirect = input.kind === "quyetDinhGianTiep";
  const khoanStyle = styleIdFromName(`${input.stylePrefix}Khoan`);
  const diemStyle = styleIdFromName(`${input.stylePrefix}Diem`);
  for (let i = 0; i < input.articleCount; i += 1) {
    parts.push(i === 0 ? legalParagraph("dieu", body, indirect ? INDIRECT_LEAD : "", MARKER_RUN) : legalParagraph("dieu", body, ""));
    for (let k = 0; k < (input.khoanPerDieu ?? 0); k += 1) {
      parts.push(legalParagraph("khoan", khoanStyle, ""));
      for (let d = 0; d < (input.diemPerKhoan ?? 0); d += 1) parts.push(diemParagraph(diemStyle, d));
    }
  }
  return parts.join("");
}

/** Dựng thông số từ Tuỳ chọn nhanh: cỡ chữ (căn cứ kẹp vào khoảng của thành phần "Căn cứ") và thụt dòng đầu. */
export function decisionInputFromOptions(
  profile: StandardProfile,
  values: FormatValues,
  text: { kind: DecisionKind; authority: string; basis: readonly string[]; articleCount: number; khoanPerDieu?: number; diemPerKhoan?: number },
): DecisionInput {
  const num = (id: string, fallback: number): number => (typeof values[id] === "number" ? (values[id] as number) : fallback);
  const size = num("fontSize", profile.styles.defaults.fontSizePt);
  const basis = profile.elements.canCu.sizePt;
  return {
    ...text,
    stylePrefix: profile.styles.prefix,
    sizes: { authority: size, basis: Math.min(basis.max, Math.max(basis.min, size)), verdict: size },
    firstLineIndentMm: num("firstLineIndent", profile.styles.defaults.firstLineIndentMm),
  };
}

/**
 * Thân văn bản đầy đủ của nghị quyết hoặc quyết định cá biệt: khối đầu, tên loại và trích yếu (biến thể có tên loại),
 * thẩm quyền, căn cứ, các Điều, khối cuối. Điều 1 mang chữ đánh dấu chỗ nội dung.
 */
export function assembleDecisionBody(input: AssembleInput, decision: DecisionInput): string {
  if (!input.title || input.title.variant !== "named") throw new Error("Nghị quyết, quyết định phải dùng khối tên loại và trích yếu.");
  return buildHeaderBlock(input.header) + buildTitleBody(input.title) + buildDecisionBody(decision) + buildFooterBlock(input.footer);
}
