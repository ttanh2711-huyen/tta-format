import type { RecipientsRow } from "../../ooxml/headerBlock";
import { paragraph } from "../../ooxml/titleBlock";
import type { FormatValues } from "../../settings/formatOptions";
import type { StandardProfile } from "../types";
import { HEADER_MARKS } from "./rules";

/**
 * Phần thể loại, trích yếu, "Kính gửi"/"Kính trình" và căn cứ ban hành của văn bản Đảng (TD3), nằm dưới bảng
 * đầu văn bản. Ba biến thể: có thể loại (thể loại, trích yếu, 5 dấu gạch nối); công văn (không ghi tên loại,
 * dòng "về việc" và "Kính gửi" nằm trong bảng đầu văn bản); tờ trình (có thể loại và "Kính trình").
 */

export type PartyTitleVariant = "congVan" | "coTenLoai" | "toTrinh";

/** Thụt treo (mm) để dòng xuống hàng thẳng hàng sau dấu gạch đầu dòng. */
const DASH_HANGING_MM = 4;

export interface PartyTitleInput {
  variant: PartyTitleVariant;
  /** Tên thể loại (biến thể có thể loại và tờ trình). */
  typeName?: string;
  /** Trích yếu: nội dung sau "về việc" của công văn, hoặc dòng trích yếu dưới thể loại. */
  subject: string;
  /** Nơi nhận của "Kính gửi" (công văn) hoặc "Kính trình" (tờ trình). */
  recipients: readonly string[];
  /** Căn cứ ban hành, mỗi phần tử một căn cứ. */
  basis: readonly string[];
  sizes: { theLoai: number; trichYeu: number; trichYeuCongVan: number; kinhGui: number };
  /** Giãn dòng Exactly (pt) của các dòng danh sách, theo giãn dòng nội dung. */
  lineExactPt: number;
  textWidthMm: number;
}

function clean(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

/** Danh sách gạch đầu dòng: bỏ dấu cũ, ";" cuối mỗi dòng, "." cuối dòng cuối; dòng trống bị bỏ. */
export function dashList(items: readonly string[]): string[] {
  const entries = items.map((item) => clean(item).replace(/^[-–—]\s*/, "").replace(/[;.,]+$/, "").trim()).filter((item) => item !== "");
  return entries.map((entry, index) => `- ${entry}${index === entries.length - 1 ? "." : ";"}`);
}

function recipientsOf(input: PartyTitleInput): string[] {
  return input.recipients.map(clean).filter((item) => item !== "");
}

const verbOf = (input: PartyTitleInput): string => (input.variant === "toTrinh" ? "Kính trình" : "Kính gửi");

export function validatePartyTitle(input: PartyTitleInput): string[] {
  const errors: string[] = [];
  if (input.variant !== "congVan" && clean(input.typeName ?? "") === "") errors.push("Hãy nhập tên thể loại văn bản.");
  if (clean(input.subject) === "") errors.push("Hãy nhập trích yếu nội dung.");
  if (input.variant !== "coTenLoai" && recipientsOf(input).length === 0) errors.push(`Hãy nhập ít nhất một nơi nhận (${verbOf(input)}).`);
  return errors;
}

/** Dòng "về việc ..." của công văn (đặt trong bảng đầu văn bản); các biến thể khác không có. */
export function partySubjectRow(input: PartyTitleInput): { text: string; sizePt: number } | undefined {
  if (input.variant !== "congVan") return undefined;
  const subject = clean(input.subject).replace(/^(?:v\/v|về việc)\s*/i, "");
  return { text: `về việc ${subject}`, sizePt: input.sizes.trichYeuCongVan };
}

/** Hàng "Kính gửi" của công văn trong bảng đầu văn bản: nhãn nghiêng canh phải ở ô trái, nơi nhận canh trái ở ô phải. */
export function partyRecipientsRow(input: PartyTitleInput): RecipientsRow | undefined {
  if (input.variant !== "congVan") return undefined;
  const names = recipientsOf(input);
  // Một nơi nhận ghi nguyên tên, cùng dòng với nhãn; từ hai nơi nhận trở lên thì gạch đầu dòng.
  return { label: "Kính gửi:", italicLabel: true, lines: names.length === 1 ? names : dashList(names), sizePt: input.sizes.kinhGui };
}

function typeAndSubject(input: PartyTitleInput): string[] {
  const { sizes } = input;
  return [
    paragraph(clean(input.typeName ?? "").toUpperCase(), { sizePt: sizes.theLoai, bold: true, beforePt: 12 }),
    paragraph(clean(input.subject), { sizePt: sizes.trichYeu, bold: true }),
    paragraph("-".repeat(HEADER_MARKS.trichYeuDashCount), { sizePt: sizes.trichYeu, afterPt: 6 }),
  ];
}

function recipientList(input: PartyTitleInput): string[] {
  const size = input.sizes.kinhGui;
  const lines = dashList(recipientsOf(input));
  return [
    paragraph(`${verbOf(input)}:`, { sizePt: size, italic: true, beforePt: 6, exactPt: input.lineExactPt }),
    ...lines.map((line) => paragraph(line, { sizePt: size, align: "left", leftMm: 0, hangingMm: DASH_HANGING_MM, exactPt: input.lineExactPt })),
  ];
}

function basisList(input: PartyTitleInput): string[] {
  return dashList(input.basis).map((line) => paragraph(line, { sizePt: input.sizes.kinhGui, align: "both", leftMm: 0, hangingMm: DASH_HANGING_MM, exactPt: input.lineExactPt }));
}

/** XML thân văn bản (các đoạn) nằm ngay sau bảng đầu văn bản. */
export function buildPartyTitleBody(input: PartyTitleInput): string {
  const problems = validatePartyTitle(input);
  if (problems.length > 0) throw new Error(problems.join(" "));
  const parts = input.variant === "congVan" ? [] : typeAndSubject(input);
  if (input.variant === "toTrinh") parts.push(...recipientList(input));
  parts.push(...basisList(input));
  return parts.join("");
}

/** Dựng thông số từ Tuỳ chọn nhanh: cỡ chữ, giãn dòng Exactly và bề rộng vùng chữ theo lề đang chọn. */
export function partyTitleFromOptions(
  profile: StandardProfile,
  values: FormatValues,
  text: { variant: PartyTitleVariant; typeName?: string; subject: string; recipients: readonly string[]; basis: readonly string[] },
): PartyTitleInput {
  const num = (id: string, fallback: number): number => (typeof values[id] === "number" ? (values[id] as number) : fallback);
  const fixed = profile.pageDefaults.marginsMm;
  return {
    ...text,
    sizes: {
      theLoai: num("size.theLoai", 16),
      trichYeu: num("size.trichYeu", 14),
      trichYeuCongVan: num("size.trichYeuCongVan", 12),
      kinhGui: num("fontSize", 14),
    },
    lineExactPt: profile.content.lineSpacing.kind === "exactly" ? num("lineSpacing", 20) : 20,
    textWidthMm: profile.page.paper.widthMm - num("marginLeft", fixed.left) - num("marginRight", fixed.right),
  };
}
