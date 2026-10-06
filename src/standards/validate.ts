import type { MarginSide, StandardProfile } from "./types";

/** Sai số cho phép (mm) khi so lề đã qua đổi đơn vị, tránh báo sai do làm tròn. */
const MARGIN_TOLERANCE_MM = 0.01;
const SPACING_TOLERANCE = 0.001;

/**
 * Lề (mm) có hợp lệ với chuẩn không. `duplexBack` là mặt sau khi in hai mặt; chuẩn không
 * quy định riêng cho mặt sau thì dùng lề mặt trước.
 */
export function isMarginValid(
  profile: StandardProfile,
  side: MarginSide,
  mm: number,
  options: { duplexBack?: boolean } = {},
): boolean {
  if (!Number.isFinite(mm)) {
    return false;
  }
  const margins =
    options.duplexBack && profile.page.duplexBackMarginsMm
      ? profile.page.duplexBackMarginsMm
      : profile.page.marginsMm;
  const { min, max } = margins[side];
  return mm >= min - MARGIN_TOLERANCE_MM && mm <= max + MARGIN_TOLERANCE_MM;
}

/** Cỡ chữ (pt) của một thành phần thể thức có nằm trong khoảng cho phép không. */
export function isFontSizeValid(profile: StandardProfile, elementId: string, pt: number): boolean {
  const element = profile.elements[elementId];
  if (!element) {
    throw new Error(`Chuẩn "${profile.name}" không có thành phần "${elementId}".`);
  }
  return Number.isFinite(pt) && pt >= element.sizePt.min && pt <= element.sizePt.max;
}

export type LineSpacingValue = { kind: "multiple"; lines: number } | { kind: "exactly"; pt: number };

/** Giãn dòng có hợp lệ không. Khác kiểu giãn dòng với quy định của chuẩn thì không hợp lệ. */
export function isLineSpacingValid(profile: StandardProfile, value: LineSpacingValue): boolean {
  const rule = profile.content.lineSpacing;
  if (rule.kind === "multiple" && value.kind === "multiple") {
    return (
      Number.isFinite(value.lines) &&
      value.lines >= rule.min - SPACING_TOLERANCE &&
      value.lines <= rule.max + SPACING_TOLERANCE
    );
  }
  if (rule.kind === "exactly" && value.kind === "exactly") {
    return (
      Number.isFinite(value.pt) &&
      value.pt >= rule.minPt - SPACING_TOLERANCE &&
      value.pt <= rule.maxPt + SPACING_TOLERANCE
    );
  }
  return false;
}

/** Thụt dòng đầu (mm) có khớp một trong các giá trị chuẩn cho phép, trong sai số không. */
export function isFirstLineIndentValid(profile: StandardProfile, mm: number): boolean {
  if (!Number.isFinite(mm)) {
    return false;
  }
  const { firstLineIndentMm, firstLineIndentToleranceMm } = profile.content;
  return firstLineIndentMm.some((allowed) => Math.abs(mm - allowed) <= firstLineIndentToleranceMm + 1e-9);
}
