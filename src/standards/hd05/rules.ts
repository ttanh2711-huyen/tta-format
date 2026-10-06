/**
 * Quy tắc thể thức và kỹ thuật trình bày văn bản của Đảng theo Hướng dẫn số 05-HD/VPTW
 * ngày 27/5/2026 của Văn phòng Trung ương Đảng (thay Hướng dẫn 36-HD/VPTW).
 *
 * Mỗi giá trị trích dẫn vị trí: "Phụ lục 2, mục N" là bảng cỡ chữ và kiểu chữ; "Phần II, ..."
 * là phần hướng dẫn cụ thể. Điểm còn mơ hồ ghi ở docs/hd05-ambiguities.md.
 */

import type {
  Casing,
  ContentLayout,
  ElementRule as GenericElementRule,
  FontRules,
  FontStyle,
  PageNumberRules,
  PageRules,
  SignAuthority,
  UrgencyMark,
} from "../types";

export type ElementId =
  | "tieuDe"
  | "coQuanCapTren"
  | "coQuanBanHanh"
  | "soKyHieu"
  | "diaDanhNgayThang"
  | "theLoai"
  | "trichYeu"
  | "trichYeuCongVan"
  | "noiDung"
  | "phanChuong"
  | "tenPhanChuong"
  | "mucSo"
  | "tenMuc"
  | "dieu"
  | "khoan"
  | "diem"
  | "quyenHan"
  | "chucVu"
  | "hoTen"
  | "kinhGuiTu"
  | "kinhGuiDanhSach"
  | "noiNhanTu"
  | "noiNhanDanhSach"
  | "dauKhan"
  | "phamViLuuHanh"
  | "duThao"
  | "kyHieuNguoiSoan"
  | "thongTinLienHe"
  | "soTrang";

export type ElementRule = GenericElementRule<ElementId>;

const SRC = "HD 05, Phụ lục 2, mục";

function element(
  id: ElementId,
  label: string,
  casing: Casing,
  min: number,
  max: number,
  style: FontStyle,
  bold: boolean,
  source: string,
  underline = false,
): ElementRule {
  const rule: ElementRule = {
    id,
    label,
    casing,
    sizePt: { min, max },
    style,
    bold,
    source: `${SRC} ${source}`,
  };
  if (underline) {
    rule.underline = true;
  }
  return rule;
}

export const ELEMENTS: Record<ElementId, ElementRule> = {
  tieuDe: element("tieuDe", "Tiêu đề \"Đảng Cộng sản Việt Nam\"", "upper", 15, 15, "upright", true, "1"),
  coQuanCapTren: element("coQuanCapTren", "Tên cơ quan cấp trên trực tiếp", "upper", 14, 14, "upright", false, "2"),
  coQuanBanHanh: element("coQuanBanHanh", "Tên cơ quan ban hành văn bản", "upper", 14, 14, "upright", true, "2"),
  soKyHieu: element("soKyHieu", "Số và ký hiệu văn bản", "lower", 14, 14, "upright", false, "3"),
  diaDanhNgayThang: element("diaDanhNgayThang", "Địa danh và ngày, tháng, năm ban hành", "lower", 14, 14, "italic", false, "4"),
  theLoai: element("theLoai", "Thể loại văn bản", "upper", 15, 16, "upright", true, "5"),
  trichYeu: element("trichYeu", "Trích yếu nội dung văn bản", "lower", 14, 15, "upright", true, "5"),
  trichYeuCongVan: element("trichYeuCongVan", "Trích yếu nội dung thể loại công văn", "lower", 12, 12, "italic", false, "5"),
  noiDung: element("noiDung", "Nội dung văn bản", "lower", 14, 15, "upright", false, "6"),
  phanChuong: element("phanChuong", "Từ \"Phần\", \"Chương\" và số thứ tự", "lower", 14, 15, "upright", true, "6"),
  tenPhanChuong: element("tenPhanChuong", "Tên phần, chương", "upper", 14, 15, "upright", true, "6"),
  mucSo: element("mucSo", "Từ \"Mục\" và số thứ tự", "lower", 14, 15, "upright", true, "6"),
  tenMuc: element("tenMuc", "Tên mục", "upper", 14, 15, "upright", true, "6"),
  dieu: element("dieu", "Từ \"Điều\", số thứ tự và tên điều", "lower", 14, 15, "upright", true, "6"),
  khoan: element("khoan", "Khoản", "lower", 14, 15, "upright", false, "6"),
  diem: element("diem", "Điểm", "lower", 14, 15, "upright", false, "6"),
  quyenHan: element("quyenHan", "Quyền hạn ký văn bản", "upper", 14, 14, "upright", true, "7"),
  chucVu: element("chucVu", "Chức vụ của người ký văn bản", "upper", 14, 14, "upright", false, "7"),
  hoTen: element("hoTen", "Họ tên của người ký văn bản", "lower", 14, 14, "upright", true, "7"),
  kinhGuiTu: element("kinhGuiTu", "Từ \"Kính gửi\", \"Kính trình\"", "lower", 14, 14, "italic", false, "9"),
  kinhGuiDanhSach: element("kinhGuiDanhSach", "Tên tổ chức, cơ quan, cá nhân sau \"Kính gửi\", \"Kính trình\"", "lower", 14, 15, "upright", false, "9"),
  noiNhanTu: element("noiNhanTu", "Từ \"Nơi nhận\"", "lower", 14, 14, "upright", false, "9", true),
  noiNhanDanhSach: element("noiNhanDanhSach", "Tên tổ chức, cơ quan, cá nhân sau \"Nơi nhận\"", "lower", 12, 12, "upright", false, "9"),
  dauKhan: element("dauKhan", "Dấu chỉ mức độ khẩn", "upper", 14, 14, "upright", true, "11"),
  phamViLuuHanh: element("phamViLuuHanh", "Chỉ dẫn phạm vi lưu hành", "upper", 12, 12, "upright", true, "12"),
  duThao: element("duThao", "Chỉ dẫn dự thảo văn bản", "lower", 14, 14, "italic", false, "12"),
  kyHieuNguoiSoan: element("kyHieuNguoiSoan", "Ký hiệu người soạn thảo và số lượng bản phát hành", "lower", 11, 11, "upright", false, "13"),
  thongTinLienHe: element("thongTinLienHe", "Thông tin liên hệ của cơ quan ban hành", "lower", 10, 10, "upright", false, "14"),
  soTrang: element("soTrang", "Số trang văn bản, phụ lục", "lower", 14, 14, "upright", false, "15"),
};

const P2 = "HD 05, Phần II, Mục IV";

export const PAGE: PageRules = {
  paper: { name: "A4", widthMm: 210, heightMm: 297, source: `${P2}.3` },
  marginsMm: {
    top: { min: 20, max: 20, source: `${P2}.4.1.1` },
    bottom: { min: 20, max: 20, source: `${P2}.4.1.2` },
    left: { min: 30, max: 30, source: `${P2}.4.1.3` },
    right: { min: 15, max: 15, source: `${P2}.4.1.4` },
  },
  duplexBackMarginsMm: {
    top: { min: 20, max: 20, source: `${P2}.4.2.1` },
    bottom: { min: 20, max: 20, source: `${P2}.4.2.2` },
    left: { min: 15, max: 15, source: `${P2}.4.2.3` },
    right: { min: 30, max: 30, source: `${P2}.4.2.4` },
  },
  landscapeAllowedForTables: true, // Mục IV.4.3
};

export const FONT: FontRules = {
  name: "Times New Roman",
  encoding: "Unicode TCVN 6909:2001",
  color: "#000000",
  source: `${P2}.1`,
};

export const PAGE_NUMBER: PageNumberRules = {
  sizePt: { min: 14, max: 14 }, // Phụ lục 2, mục 15
  align: "center",
  position: "top-margin",
  distanceFromTopEdgeMm: 10,
  hideOnFirstPage: true,
  appendixNumberedSeparately: true,
  source: `${P2}.5`,
};

export const CONTENT_LAYOUT: ContentLayout = {
  align: "justify",
  /** "Chữ đầu dòng lùi vào khoảng 10 mm" nên cho sai số 1 mm. */
  firstLineIndentMm: [10],
  firstLineIndentToleranceMm: 1,
  paragraphSpacingMinPt: 6,
  lineSpacing: { kind: "exactly", minPt: 18, maxPt: 22 },
  endsWithPeriod: true,
  source: "HD 05, Phần II, Mục I.6.2.2",
};

/** Dấu hiệu nhận biết và quy cách riêng của chuẩn Đảng ở khối đầu văn bản. */
export const HEADER_MARKS = {
  /** Tiêu đề ở góc phải, dòng đầu; dưới có đường kẻ ngang nét liền dài bằng độ dài tiêu đề. */
  identifyingTitle: "ĐẢNG CỘNG SẢN VIỆT NAM",
  /** Dấu sao dưới tên cơ quan ban hành, ngăn cách với số và ký hiệu. */
  agencyStar: "*",
  /** Số dấu gạch nối dưới trích yếu, ngăn cách với nội dung. */
  trichYeuDashCount: 5,
  /** Từ "Số" không có dấu hai chấm: "Số 05-HD/VPTW". */
  soPrefix: "Số",
  source: "HD 05, Phần II, Mục I.1.2, I.2.2.1, I.3.2, I.5.2.1",
};

export const SIGN_AUTHORITY: readonly SignAuthority[] = [
  { abbr: "T/M", meaning: "thay mặt (đại hội, cấp ủy, ủy ban kiểm tra, hội đồng)" },
  { abbr: "K/T", meaning: "ký thay (cấp phó ký thay cấp trưởng)" },
  { abbr: "T/L", meaning: "thừa lệnh" },
  { abbr: "Q.", meaning: "quyền cấp trưởng" },
];

/** Dấu chỉ mức độ khẩn đặt dưới số và ký hiệu (công văn: dưới trích yếu). HD 05 không nêu kích thước. */
export const URGENCY_MARKS: readonly UrgencyMark[] = [
  { text: "KHẨN" },
  { text: "THƯỢNG KHẨN" },
  { text: "HỎA TỐC" },
];
