import { formatSo } from "../dateNumber";
import { assertAbbreviation, assertAgencyAbbreviations } from "../soKyHieu";
import type { StandardProfile } from "../types";
import { DOC_TYPES } from "./docTypes";
import { STYLE_PLAN } from "./styles";
import {
  CONTENT_LAYOUT,
  ELEMENTS,
  FONT,
  LINE_RULES,
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
  pageDefaults: {
    marginsMm: { top: 20, bottom: 20, left: 30, right: 20 },
    headerDistanceMm: 12.5,
    footerDistanceMm: 5.3,
    source: "File mẫu công văn của người dùng (lề 20/20/30/20 mm, header 12,5 mm, footer khoảng 5,3 mm); lề nằm trong khoảng của NĐ 30",
  },
  headerKind: "quocHieu",
  agencyLineRatio: LINE_RULES.tenCoQuanBanHanh.lengthRatio,
  titleLineRatio: LINE_RULES.trichYeu.lengthRatio,
  defaultPresetName: "Theo mẫu công văn của tôi",
  styles: STYLE_PLAN,
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
