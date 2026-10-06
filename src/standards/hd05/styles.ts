import type { StylePlan } from "../types";

/**
 * Bộ style của chuẩn Đảng. Giá trị mặc định là đề xuất trong khoảng của HD 05 (kế hoạch, Q13):
 * cỡ chữ 14, giãn dòng Exactly 20 pt, thụt dòng đầu 10 mm, giãn đoạn 6 pt.
 */
export const STYLE_PLAN: StylePlan = {
  prefix: "HD05_",
  defaults: {
    fontSizePt: 14,
    spaceBeforePt: 6,
    spaceAfterPt: 6,
    firstLineIndentMm: 10,
    lineSpacing: { kind: "exactly", pt: 20 },
  },
  elements: [
    { elementId: "noiDung", alignment: "Justified", indent: true },
    { elementId: "kinhGuiTu", alignment: "Centered" },
    { elementId: "kinhGuiDanhSach", alignment: "Left" },
    { elementId: "phanChuong", alignment: "Centered", keepWithNext: true },
    { elementId: "tenPhanChuong", alignment: "Centered", keepWithNext: true },
    { elementId: "mucSo", alignment: "Centered", keepWithNext: true },
    { elementId: "tenMuc", alignment: "Centered", keepWithNext: true },
    { elementId: "dieu", alignment: "Justified", indent: true, keepWithNext: true },
    { elementId: "khoan", alignment: "Justified", indent: true },
    { elementId: "diem", alignment: "Justified", indent: true },
  ],
  table: { fontSizePt: 14, spaceBeforePt: 3, spaceAfterPt: 3 },
};
