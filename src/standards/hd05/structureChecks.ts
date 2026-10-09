import { checkStructureTitles, type StructureTitle } from "../../checker/structureRules";
import type { CheckDocument, Issue } from "../../checker/types";
import type { StandardProfile } from "../types";

/** Cấu trúc Phần/Chương, Mục, Điều của chuẩn Đảng (TD6): tên Phần/Chương và tên Mục in hoa đậm, Điều đậm. */
export function checkPartyStructure(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const prefix = profile.styles.prefix;
  const titles: StructureTitle[] = [
    { style: `${prefix}TenPhanChuong`, rule: "chuong", label: "Tên Phần, Chương", elementId: "tenPhanChuong" },
    { style: `${prefix}TenMuc`, rule: "muc", label: "Tên Mục", elementId: "tenMuc" },
  ];
  return checkStructureTitles(doc, profile, titles, `${prefix}Dieu`);
}
