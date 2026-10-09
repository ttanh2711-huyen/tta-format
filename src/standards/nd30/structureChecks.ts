import { checkStructureTitles, type StructureTitle } from "../../checker/structureRules";
import type { CheckDocument, Issue } from "../../checker/types";
import type { StandardProfile } from "../types";

/** Cấu trúc Chương, Mục, Điều và tên phụ lục của chuẩn Chính quyền (T18, T21). */
export function checkNd30Structure(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const prefix = profile.styles.prefix;
  const titles: StructureTitle[] = [
    { style: `${prefix}TieuDePhanChuong`, rule: "chuong", label: "Tên Chương", elementId: "tieuDePhanChuong" },
    { style: `${prefix}TieuDeMucTieuMuc`, rule: "muc", label: "Tên Mục", elementId: "tieuDeMucTieuMuc" },
    { style: `${prefix}PhuLucTieuDe`, rule: "phuLuc", label: "Tên phụ lục", elementId: "phuLucTieuDe" },
  ];
  return checkStructureTitles(doc, profile, titles, `${prefix}Dieu`);
}
