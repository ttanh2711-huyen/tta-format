import { formatSo } from "../dateNumber";
import { assertAbbreviation, assertAgencyAbbreviations } from "../soKyHieu";
import type { StandardProfile } from "../types";
import { AGENCY_CODES } from "./agencyCodes";
import { DOC_TYPES } from "./docTypes";
import { STYLE_PLAN } from "./styles";
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
  pageDefaults: {
    marginsMm: { top: 20, bottom: 20, left: 30, right: 15 },
    headerDistanceMm: 10,
    footerDistanceMm: 12.5,
    source:
      "HD 05 Phần II Mục IV.4.1 (lề) và IV.5.1 (số trang cách mép trên 10 mm); khoảng cách footer HD 05 không quy định, lấy mặc định của Word 12,5 mm",
  },
  headerKind: "dang",
  agencyCodes: AGENCY_CODES,
  defaultPresetName: "Theo HD 05 (mặc định đề xuất)",
  styles: STYLE_PLAN,
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
