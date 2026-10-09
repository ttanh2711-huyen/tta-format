import { isMarginValid } from "../standards/validate";
import type { MarginSide, StandardProfile } from "../standards/types";
import type { Issue, PageInfo } from "./types";

const PAPER_TOLERANCE_MM = 0.5;
const HEADER_DISTANCE_TOLERANCE_MM = 0.5;

const SIDE_NAME: Record<MarginSide, string> = { top: "trên", bottom: "dưới", left: "trái", right: "phải" };

function vn(n: number): string {
  return String(Math.round(n * 100) / 100).replace(".", ",");
}

function rangeText(min: number, max: number): string {
  return min === max ? `${vn(min)} mm` : `${vn(min)}–${vn(max)} mm`;
}

/** Kiểm tra khổ giấy, hướng trang, lề và khoảng cách số trang của một phần theo chuẩn. */
export function checkPage(page: PageInfo, section: number, profile: StandardProfile): Issue[] {
  const issues: Issue[] = [];
  const location = { kind: "page", section } as const;
  const { widthMm, heightMm } = profile.page.paper;
  const near = (a: number, b: number) => Math.abs(a - b) <= PAPER_TOLERANCE_MM;
  const portrait = near(page.widthMm, widthMm) && near(page.heightMm, heightMm);
  const landscape = near(page.widthMm, heightMm) && near(page.heightMm, widthMm);

  if (!portrait && !landscape) {
    issues.push({
      ruleId: "page.paper",
      severity: "error",
      message: `Khổ giấy ${vn(page.widthMm)} × ${vn(page.heightMm)} mm không đúng; chuẩn ${profile.name} dùng ${profile.page.paper.name} (${widthMm} × ${heightMm} mm).`,
      reference: profile.page.paper.source,
      actual: `${vn(page.widthMm)} × ${vn(page.heightMm)} mm`,
      expected: `${widthMm} × ${heightMm} mm`,
      location,
      fix: { kind: "page", section, paper: "A4", label: `Đặt khổ giấy ${profile.page.paper.name}, trang đứng` },
    });
  }
  if (landscape) {
    issues.push({
      ruleId: "page.orientation",
      severity: profile.page.landscapeAllowedForTables ? "warning" : "error",
      message: profile.page.landscapeAllowedForTables
        ? "Trang ngang chỉ dùng cho nội dung có bảng, biểu; hãy kiểm tra phần này."
        : `Chuẩn ${profile.name} yêu cầu trang đứng.`,
      reference: profile.page.paper.source,
      actual: "trang ngang",
      expected: "trang đứng",
      location,
    });
    return issues;
  }

  for (const side of Object.keys(SIDE_NAME) as MarginSide[]) {
    const value = page.marginsMm[side];
    if (isMarginValid(profile, side, value)) continue;
    const range = profile.page.marginsMm[side];
    const target = Math.min(range.max, Math.max(range.min, profile.pageDefaults.marginsMm[side]));
    issues.push({
      ruleId: `page.margin.${side}`,
      severity: "error",
      message: `Lề ${SIDE_NAME[side]} ${vn(value)} mm không đúng; chuẩn ${profile.name} ${range.min === range.max ? "quy định đúng" : "cho phép"} ${rangeText(range.min, range.max)}.`,
      reference: range.source,
      actual: `${vn(value)} mm`,
      expected: rangeText(range.min, range.max),
      location,
      fix: { kind: "page", section, marginsMm: { [side]: target }, label: `Đặt lề ${SIDE_NAME[side]} ${vn(target)} mm` },
    });
  }

  const wanted = profile.pageNumber.distanceFromTopEdgeMm;
  if (wanted !== undefined && Math.abs(page.headerDistanceMm - wanted) > HEADER_DISTANCE_TOLERANCE_MM) {
    issues.push({
      ruleId: "page.headerDistance",
      severity: "warning",
      message: `Khoảng cách header ${vn(page.headerDistanceMm)} mm; số trang cần cách mép trên giấy ${vn(wanted)} mm.`,
      reference: profile.pageNumber.source,
      actual: `${vn(page.headerDistanceMm)} mm`,
      expected: `${vn(wanted)} mm`,
      location,
      fix: { kind: "page", section, headerDistanceMm: wanted, label: `Đặt khoảng cách header ${vn(wanted)} mm` },
    });
  }
  return issues;
}
