import type { StylePlan } from "../types";

/** Bộ style của chuẩn Chính quyền; giá trị mặc định theo file mẫu công văn (mục 6A của kế hoạch). */
export const STYLE_PLAN: StylePlan = {
  prefix: "ND30_",
  defaults: {
    fontSizePt: 14,
    spaceBeforePt: 6,
    spaceAfterPt: 6,
    firstLineIndentMm: 12.7,
    lineSpacing: { kind: "multiple", lines: 1.17 },
  },
  elements: [
    { elementId: "noiDung", alignment: "Justified", indent: true },
    { elementId: "kinhGui", alignment: "Centered", spacingPt: { before: 12, after: 12 } },
    { elementId: "canCu", alignment: "Justified", indent: true },
    { elementId: "phanChuong", alignment: "Centered", keepWithNext: true },
    { elementId: "tieuDePhanChuong", alignment: "Centered", keepWithNext: true },
    { elementId: "mucTieuMuc", alignment: "Centered", keepWithNext: true },
    { elementId: "tieuDeMucTieuMuc", alignment: "Centered", keepWithNext: true },
    { elementId: "dieu", alignment: "Justified", indent: true, keepWithNext: true },
    { elementId: "khoan", alignment: "Justified", indent: true },
    { elementId: "khoanCoTieuDe", alignment: "Justified", indent: true, keepWithNext: true },
    { elementId: "diem", alignment: "Justified", indent: true },
    { elementId: "phuLucSo", alignment: "Centered", keepWithNext: true },
    { elementId: "phuLucTieuDe", alignment: "Centered", keepWithNext: true },
    { elementId: "phuLucChiDan", alignment: "Centered" },
  ],
  table: { fontSizePt: 14, spaceBeforePt: 3, spaceAfterPt: 3 },
};
