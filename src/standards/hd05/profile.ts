import { formatSo } from "../dateNumber";
import { assertAbbreviation, assertAgencyAbbreviations } from "../soKyHieu";
import type { StandardProfile } from "../types";
import { DOC_TYPES } from "./docTypes";
import {
  CONTENT_LAYOUT,
  ELEMENTS,
  FONT,
  HEADER_MARKS,
  PAGE,
  PAGE_NUMBER,
  SIGN_AUTHORITY,
  URGENCY_MARKS,
} from "./rules";

export const HD05: StandardProfile = {
  id: "hd05",
  name: "Đảng (HD 05)",
  legalBasis:
    "Hướng dẫn số 05-HD/VPTW ngày 27/5/2026 của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng",
  effectiveDate: "2026-05-27",
  identifyingTitle: HEADER_MARKS.identifyingTitle,
  page: PAGE,
  font: FONT,
  pageNumber: PAGE_NUMBER,
  content: CONTENT_LAYOUT,
  elements: ELEMENTS,
  docTypes: DOC_TYPES,
  signAuthority: SIGN_AUTHORITY,
  urgencyMarks: URGENCY_MARKS,

  /** "Số 05-HD/VPTW": không có dấu hai chấm, số nối với tên loại bằng "-", tên loại nối với cơ quan bằng "/". */
  formatSoKyHieu(so, typeAbbr, agencyAbbrs) {
    const number = formatSo(so);
    if (typeAbbr === null) {
      throw new RangeError("Ký hiệu văn bản của Đảng cần chữ viết tắt tên loại văn bản.");
    }
    assertAbbreviation(typeAbbr, "Chữ viết tắt tên loại văn bản");
    assertAgencyAbbreviations(agencyAbbrs);
    return `${HEADER_MARKS.soPrefix} ${number}-${typeAbbr}/${agencyAbbrs.join("-")}`;
  },
};
