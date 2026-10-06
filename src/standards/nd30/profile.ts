import { formatSo } from "../dateNumber";
import { assertAbbreviation, assertAgencyAbbreviations } from "../soKyHieu";
import type { StandardProfile } from "../types";
import { DOC_TYPES } from "./docTypes";
import {
  CONTENT_LAYOUT,
  ELEMENTS,
  FONT,
  PAGE,
  PAGE_NUMBER,
  SIGN_AUTHORITY,
  URGENCY_MARKS,
} from "./rules";

export const ND30: StandardProfile = {
  id: "nd30",
  name: "Chính quyền (NĐ 30)",
  legalBasis: "Nghị định số 30/2020/NĐ-CP ngày 05/3/2020 của Chính phủ về công tác văn thư",
  effectiveDate: "2020-03-05",
  identifyingTitle: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
  page: PAGE,
  font: FONT,
  pageNumber: PAGE_NUMBER,
  content: CONTENT_LAYOUT,
  elements: ELEMENTS,
  docTypes: DOC_TYPES,
  signAuthority: SIGN_AUTHORITY,
  urgencyMarks: URGENCY_MARKS,

  /** "Số: 15/QĐ-BNV"; công văn không có chữ viết tắt tên loại: "Số: 05/BNV-VP". */
  formatSoKyHieu(so, typeAbbr, agencyAbbrs) {
    const number = formatSo(so);
    assertAgencyAbbreviations(agencyAbbrs);
    if (typeAbbr !== null) {
      assertAbbreviation(typeAbbr, "Chữ viết tắt tên loại văn bản");
    }
    const parts = typeAbbr === null ? agencyAbbrs : [typeAbbr, ...agencyAbbrs];
    return `Số: ${number}/${parts.join("-")}`;
  },
};
