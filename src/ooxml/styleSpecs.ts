import { mmToPt } from "../standards/units";
import type { LineSpacingSetting, ParagraphAlignment, StandardProfile, StyleOptions } from "../standards/types";

/** Mô tả một style đoạn, độc lập với cách cài vào Word (JSON hay OOXML). */
export interface StyleSpec {
  /** Tên hiển thị trong Word, ví dụ "ND30_NoiDung". */
  name: string;
  /** Mã thành phần thể thức; "bangDuLieu" cho style chữ trong bảng. */
  elementId: string;
  label: string;
  font: { name: string; sizePt: number; bold: boolean; italic: boolean; underline: boolean; color: string };
  paragraph: {
    alignment: ParagraphAlignment;
    firstLineIndentPt: number;
    spaceBeforePt: number;
    spaceAfterPt: number;
    lineSpacing: LineSpacingSetting;
    keepWithNext: boolean;
  };
}

export const TABLE_ELEMENT_ID = "bangDuLieu";

export function defaultStyleOptions(profile: StandardProfile): StyleOptions {
  const { defaults } = profile.styles;
  return { ...defaults, lineSpacing: { ...defaults.lineSpacing } };
}

function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

function pascal(id: string): string {
  return id.charAt(0).toUpperCase() + id.slice(1);
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

/**
 * Dựng danh sách style của một chuẩn từ quy tắc và tuỳ chọn. Cỡ chữ được kéo về khoảng cho phép
 * của từng thành phần; giãn đoạn, thụt dòng và giãn dòng chung chỉ áp dụng cho style mà kế hoạch
 * của chuẩn không đặt giá trị riêng.
 */
export function buildStyleSpecs(profile: StandardProfile, options: StyleOptions = defaultStyleOptions(profile)): StyleSpec[] {
  const { prefix, elements, table } = profile.styles;
  const specs: StyleSpec[] = elements.map((plan) => {
    const rule = profile.elements[plan.elementId];
    if (!rule) {
      throw new Error(`Chuẩn ${profile.name} không có thành phần "${plan.elementId}" để làm style.`);
    }
    return {
      name: `${prefix}${pascal(plan.elementId)}`,
      elementId: plan.elementId,
      label: rule.label,
      font: {
        name: profile.font.name,
        sizePt: clamp(options.fontSizePt, rule.sizePt.min, rule.sizePt.max),
        bold: rule.bold,
        italic: rule.style === "italic",
        underline: rule.underline === true,
        color: profile.font.color,
      },
      paragraph: {
        alignment: plan.alignment,
        firstLineIndentPt: plan.indent ? round2(mmToPt(options.firstLineIndentMm)) : 0,
        spaceBeforePt: plan.spacingPt?.before ?? options.spaceBeforePt,
        spaceAfterPt: plan.spacingPt?.after ?? options.spaceAfterPt,
        lineSpacing: options.lineSpacing,
        keepWithNext: plan.keepWithNext === true,
      },
    };
  });

  specs.push({
    name: `${prefix}BangDuLieu`,
    elementId: TABLE_ELEMENT_ID,
    label: "Chữ trong bảng biểu",
    font: { name: profile.font.name, sizePt: table.fontSizePt, bold: false, italic: false, underline: false, color: profile.font.color },
    paragraph: {
      alignment: "Left",
      firstLineIndentPt: 0,
      spaceBeforePt: table.spaceBeforePt,
      spaceAfterPt: table.spaceAfterPt,
      lineSpacing: { kind: "multiple", lines: 1 },
      keepWithNext: false,
    },
  });
  return specs;
}
