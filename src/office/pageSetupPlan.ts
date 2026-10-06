import { mmToPt } from "../standards/units";
import { isMarginValid } from "../standards/validate";
import type { MarginSide, StandardProfile } from "../standards/types";

export interface PageSetupInput {
  /** Lề người dùng nhập (mm); cạnh nào không có thì dùng mặc định của chuẩn. */
  marginsMm?: Partial<Record<MarginSide, number>>;
  /** In hai mặt: bật lề đối xứng. Chỉ chuẩn có quy định lề mặt sau mới cho phép. */
  duplex?: boolean;
}

export interface PageSetupSettings {
  paperSize: "A4";
  orientation: "Portrait";
  marginsMm: Record<MarginSide, number>;
  headerDistanceMm: number;
  footerDistanceMm: number;
  mirrorMargins: boolean;
}

export type PageSetupPlan = { ok: true; settings: PageSetupSettings } | { ok: false; errors: string[] };

const SIDES: readonly MarginSide[] = ["top", "bottom", "left", "right"];
const SIDE_NAME: Record<MarginSide, string> = { top: "trên", bottom: "dưới", left: "trái", right: "phải" };

function vn(n: number): string {
  return String(n).replace(".", ",");
}

/** Đọc số mm người dùng gõ: chấp nhận dấu chấm hoặc dấu phẩy thập phân; không hợp lệ thì NaN. */
export function parseMm(text: string): number {
  const trimmed = text.trim();
  return /^\d+([.,]\d+)?$/.test(trimmed) ? Number(trimmed.replace(",", ".")) : Number.NaN;
}

/**
 * Lập kế hoạch thiết lập trang theo chuẩn: A4 dọc, lề mặc định của chuẩn, thay bằng lề người dùng
 * nhập nếu hợp lệ. Lề ngoài khoảng của chuẩn thì **không áp dụng** và trả về danh sách lỗi.
 */
export function planPageSetup(profile: StandardProfile, input: PageSetupInput = {}): PageSetupPlan {
  const defaults = profile.pageDefaults;
  const errors: string[] = [];
  const marginsMm = { ...defaults.marginsMm };

  for (const side of SIDES) {
    const wanted = input.marginsMm?.[side];
    if (wanted === undefined) {
      continue;
    }
    const label = `Lề ${SIDE_NAME[side]}`;
    if (!Number.isFinite(wanted)) {
      errors.push(`${label} chưa nhập số hợp lệ.`);
      continue;
    }
    if (!isMarginValid(profile, side, wanted)) {
      const { min, max } = profile.page.marginsMm[side];
      const allowed = min === max ? `quy định đúng ${vn(min)} mm` : `chỉ cho phép trong khoảng ${vn(min)}–${vn(max)} mm`;
      errors.push(`${label} ${vn(wanted)} mm không hợp lệ: chuẩn ${profile.name} ${allowed}.`);
      continue;
    }
    marginsMm[side] = wanted;
  }

  const duplex = input.duplex === true;
  if (duplex && !profile.page.duplexBackMarginsMm) {
    errors.push(`Chuẩn ${profile.name} không quy định in hai mặt.`);
  }

  if (errors.length > 0) {
    return { ok: false, errors };
  }
  return {
    ok: true,
    settings: {
      paperSize: "A4",
      orientation: "Portrait",
      marginsMm,
      headerDistanceMm: defaults.headerDistanceMm,
      footerDistanceMm: defaults.footerDistanceMm,
      mirrorMargins: duplex,
    },
  };
}

/** Giá trị Word trả về khi đọc lại (đơn vị pt). */
export interface PageSetupReadback {
  widthPt: number;
  heightPt: number;
  topPt: number;
  bottomPt: number;
  leftPt: number;
  rightPt: number;
  headerDistancePt: number;
  footerDistancePt: number;
  mirrorMargins: boolean;
}

const TOLERANCE_MM = 0.1;
const PAPER_TOLERANCE_MM = 0.5;
const A4_MM = { width: 210, height: 297 };

/** So giá trị đọc lại với kế hoạch; trả về danh sách chênh lệch bằng tiếng Việt, rỗng nếu khớp. */
export function findPageSetupMismatches(expected: PageSetupSettings, actual: PageSetupReadback): string[] {
  const mismatches: string[] = [];
  const check = (label: string, actualPt: number, expectedMm: number, toleranceMm: number) => {
    const actualMm = actualPt / mmToPt(1);
    if (!Number.isFinite(actualMm) || Math.abs(actualMm - expectedMm) > toleranceMm) {
      mismatches.push(`${label}: đọc lại ${vn(Math.round(actualMm * 100) / 100)} mm, mong đợi ${vn(expectedMm)} mm`);
    }
  };
  check("chiều rộng khổ giấy", actual.widthPt, A4_MM.width, PAPER_TOLERANCE_MM);
  check("chiều cao khổ giấy", actual.heightPt, A4_MM.height, PAPER_TOLERANCE_MM);
  check("lề trên", actual.topPt, expected.marginsMm.top, TOLERANCE_MM);
  check("lề dưới", actual.bottomPt, expected.marginsMm.bottom, TOLERANCE_MM);
  check("lề trái", actual.leftPt, expected.marginsMm.left, TOLERANCE_MM);
  check("lề phải", actual.rightPt, expected.marginsMm.right, TOLERANCE_MM);
  check("khoảng cách header", actual.headerDistancePt, expected.headerDistanceMm, TOLERANCE_MM);
  check("khoảng cách footer", actual.footerDistancePt, expected.footerDistanceMm, TOLERANCE_MM);
  if (actual.mirrorMargins !== expected.mirrorMargins) {
    mismatches.push(`lề đối xứng: đọc lại ${actual.mirrorMargins ? "bật" : "tắt"}, mong đợi ${expected.mirrorMargins ? "bật" : "tắt"}`);
  }
  return mismatches;
}

export interface ApplySummary {
  kind: "ok" | "warning";
  text: string;
}

/** Câu thông báo kết quả sau khi áp dụng và đọc lại. */
export function summarizeApply(settings: PageSetupSettings, sections: number, mismatches: readonly string[]): ApplySummary {
  const { top, bottom, left, right } = settings.marginsMm;
  const margins = `lề trên ${vn(top)}, dưới ${vn(bottom)}, trái ${vn(left)}, phải ${vn(right)} mm`;
  const described = settings.mirrorMargins ? `${margins}, bật lề đối xứng,` : margins;
  if (mismatches.length === 0) {
    return { kind: "ok", text: `Đã thiết lập khổ A4, ${described} cho ${sections} phần.` };
  }
  return {
    kind: "warning",
    text: `Đã áp dụng nhưng đọc lại không khớp: ${mismatches.join("; ")}.`,
  };
}
