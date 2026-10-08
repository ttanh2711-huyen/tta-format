import type { FormatValues } from "../settings/formatOptions";
import type { StandardProfile } from "../standards/types";
import { lineParagraph } from "./headerBlock";
import { mmToTwips } from "./flatOpc";
import { textWidthMm } from "./textWidth";

/**
 * Phần tên loại, trích yếu và "Kính gửi" nằm dưới bảng đầu văn bản (T9). Hai biến thể:
 * - có tên loại (quyết định, tờ trình...): tên loại in hoa đậm, trích yếu đậm, đường kẻ dưới trích yếu;
 * - công văn: dòng "V/v ..." nằm trong bảng đầu văn bản (hàng thứ ba, xem `subjectLine`), còn ở đây là "Kính gửi".
 */

const FONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>';
const KINH_GUI = "Kính gửi:";
/** Thụt treo (mm) của dòng nơi nhận để chữ xuống dòng thẳng hàng sau dấu gạch đầu dòng. */
const DASH_HANGING_MM = 4;

export interface TitleBlockInput {
  variant: "named" | "letter";
  /** Tên loại văn bản (biến thể có tên loại). */
  typeName?: string;
  /** Trích yếu: nội dung sau "V/v" của công văn, hoặc dòng trích yếu dưới tên loại. */
  subject: string;
  /** Các nơi nhận của "Kính gửi" (công văn). */
  recipients: readonly string[];
  sizes: { tenLoai: number; trichYeu: number; kinhGui: number };
  /** Độ dài đường kẻ dưới trích yếu so với độ dài dòng chữ. */
  titleLineRatio: number;
  textWidthMm: number;
  /** Giãn dòng bội số của dòng "Kính gửi" (theo giãn dòng nội dung). */
  lineMultiple: number;
}

function clean(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

function escapeXml(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Dòng "V/v ..." của công văn; không thêm tiền tố nếu người dùng đã gõ. */
export function subjectLine(subject: string): string {
  const text = clean(subject);
  return /^v\/v\b/i.test(text) ? `V/v${text.slice(3)}` : `V/v ${text}`;
}

function recipientsOf(input: TitleBlockInput): string[] {
  return input.recipients.map((name) => clean(name).replace(/^[-–]\s*/, "").replace(/[;.,]+$/, "")).filter((name) => name !== "");
}

export function validateTitleInput(input: TitleBlockInput): string[] {
  const errors: string[] = [];
  if (input.variant === "named" && clean(input.typeName ?? "") === "") errors.push("Hãy nhập tên loại văn bản.");
  if (clean(input.subject) === "") errors.push("Hãy nhập trích yếu nội dung.");
  if (input.variant === "letter" && recipientsOf(input).length === 0) errors.push("Hãy nhập ít nhất một nơi nhận (Kính gửi).");
  return errors;
}

interface ParagraphStyle {
  sizePt: number;
  bold?: boolean;
  align?: "center" | "left";
  beforePt?: number;
  afterPt?: number;
  lineMultiple?: number;
  leftMm?: number;
  hangingMm?: number;
}

function paragraph(text: string, style: ParagraphStyle): string {
  const half = style.sizePt * 2;
  // Word ghi giãn dòng bội số theo 1/240 dòng, cắt phần lẻ (1,17 dòng là 280).
  const line = Math.floor((style.lineMultiple ?? 1) * 240 + 1e-6);
  const indent =
    style.leftMm === undefined
      ? ""
      : `<w:ind w:left="${mmToTwips(style.leftMm + (style.hangingMm ?? 0))}"${style.hangingMm ? ` w:hanging="${mmToTwips(style.hangingMm)}"` : ""}/>`;
  return (
    `<w:p><w:pPr><w:spacing w:before="${(style.beforePt ?? 0) * 20}" w:after="${(style.afterPt ?? 0) * 20}" w:line="${line}" w:lineRule="auto"/>${indent}<w:jc w:val="${style.align ?? "center"}"/></w:pPr>` +
    `<w:r><w:rPr>${FONT}${style.bold ? "<w:b/>" : ""}<w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`
  );
}

function namedBody(input: TitleBlockInput): string {
  const subject = clean(input.subject);
  const width = Math.min(textWidthMm(subject, input.sizes.trichYeu, true), input.textWidthMm);
  return [
    paragraph(clean(input.typeName ?? "").toUpperCase(), { sizePt: input.sizes.tenLoai, bold: true, beforePt: 12 }),
    paragraph(subject, { sizePt: input.sizes.trichYeu, bold: true }),
    lineParagraph(width * input.titleLineRatio, mmToTwips(input.textWidthMm)),
  ].join("");
}

function letterBody(input: TitleBlockInput): string {
  const names = recipientsOf(input);
  const size = input.sizes.kinhGui;
  const base = { sizePt: size, lineMultiple: input.lineMultiple };
  if (names.length === 1) return paragraph(`${KINH_GUI} ${names[0]}`, { ...base, beforePt: 12, afterPt: 12 });
  // Nhiều nơi nhận: mỗi nơi một dòng, gạch đầu dòng, cuối dòng chấm phẩy, dòng cuối chấm; danh sách thẳng hàng với "Kính gửi:".
  const left = Math.max(0, (input.textWidthMm - textWidthMm(KINH_GUI, size, false)) / 2);
  const list = names.map((name, index) => {
    const last = index === names.length - 1;
    return paragraph(`- ${name}${last ? "." : ";"}`, { ...base, align: "left", leftMm: left, hangingMm: DASH_HANGING_MM, afterPt: last ? 12 : 0 });
  });
  return [paragraph(KINH_GUI, { ...base, beforePt: 12 }), ...list].join("");
}

/** XML thân văn bản (các đoạn) nằm ngay sau bảng đầu văn bản. */
export function buildTitleBody(input: TitleBlockInput): string {
  const problems = validateTitleInput(input);
  if (problems.length > 0) throw new Error(problems.join(" "));
  return input.variant === "named" ? namedBody(input) : letterBody(input);
}

/** Dựng thông số tên loại, trích yếu, "Kính gửi" từ Tuỳ chọn nhanh. */
export function titleInputFromOptions(
  profile: StandardProfile,
  values: FormatValues,
  text: { variant: "named" | "letter"; typeName?: string; subject: string; recipients: readonly string[] },
): TitleBlockInput {
  const num = (id: string, fallback: number): number => (typeof values[id] === "number" ? (values[id] as number) : fallback);
  const fixed = profile.pageDefaults.marginsMm;
  return {
    ...text,
    sizes: { tenLoai: num("size.tenLoai", 14), trichYeu: num("size.trichYeu", 14), kinhGui: num("fontSize", 14) },
    titleLineRatio: num("titleLineRatio", profile.titleLineRatio?.max ?? 0.5),
    textWidthMm: profile.page.paper.widthMm - num("marginLeft", fixed.left) - num("marginRight", fixed.right),
    lineMultiple: profile.content.lineSpacing.kind === "multiple" ? num("lineSpacing", 1) : 1,
  };
}
