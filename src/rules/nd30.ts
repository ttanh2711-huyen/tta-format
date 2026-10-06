/**
 * Quy tắc thể thức và kỹ thuật trình bày văn bản hành chính theo
 * Nghị định 30/2020/NĐ-CP ngày 05/3/2020 của Chính phủ.
 *
 * Mỗi giá trị có trường `source` trích dẫn vị trí trong Phụ lục I (Phần I) hoặc
 * Phụ lục III để người rà soát đối chiếu được. Các điểm còn mơ hồ ghi ở
 * docs/nd30-ambiguities.md.
 */

export interface Range {
  min: number;
  max: number;
}

export type Casing = "upper" | "lower"; // chữ in hoa | chữ in thường
export type FontStyle = "upright" | "italic"; // kiểu chữ đứng | nghiêng

export type ElementId =
  | "quocHieu"
  | "tieuNgu"
  | "tenCoQuanChuQuan"
  | "tenCoQuanBanHanh"
  | "so"
  | "kyHieu"
  | "diaDanhNgayThang"
  | "tenLoai"
  | "trichYeu"
  | "trichYeuCongVan"
  | "canCu"
  | "noiDung"
  | "phanChuong"
  | "tieuDePhanChuong"
  | "mucTieuMuc"
  | "tieuDeMucTieuMuc"
  | "dieu"
  | "khoan"
  | "khoanCoTieuDe"
  | "diem"
  | "quyenHanChucVu"
  | "hoTenNguoiKy"
  | "kinhGui"
  | "noiNhanTieuDe"
  | "noiNhanDanhSach"
  | "phuLucSo"
  | "phuLucTieuDe"
  | "phuLucChiDan"
  | "dauKhan"
  | "phamViLuuHanh"
  | "kyHieuNguoiSoan"
  | "diaChiCoQuan";

export interface ElementRule {
  id: ElementId;
  label: string;
  casing: Casing;
  sizePt: Range;
  style: FontStyle;
  bold: boolean;
  source: string;
}

const PL1 = "Phụ lục I, Phần I, Mục";

function element(
  id: ElementId,
  label: string,
  casing: Casing,
  min: number,
  max: number,
  style: FontStyle,
  bold: boolean,
  source: string,
): ElementRule {
  return { id, label, casing, sizePt: { min, max }, style, bold, source: `${PL1} ${source}` };
}

export const ELEMENTS: Record<ElementId, ElementRule> = {
  quocHieu: element("quocHieu", "Quốc hiệu", "upper", 12, 13, "upright", true, "II.1.a"),
  tieuNgu: element("tieuNgu", "Tiêu ngữ", "lower", 13, 14, "upright", true, "II.1.b"),
  tenCoQuanChuQuan: element("tenCoQuanChuQuan", "Tên cơ quan, tổ chức chủ quản trực tiếp", "upper", 12, 13, "upright", false, "II.2.b"),
  tenCoQuanBanHanh: element("tenCoQuanBanHanh", "Tên cơ quan, tổ chức ban hành văn bản", "upper", 12, 13, "upright", true, "II.2.b"),
  so: element("so", "Từ \"Số\" của số, ký hiệu văn bản", "lower", 13, 13, "upright", false, "II.3.c"),
  kyHieu: element("kyHieu", "Ký hiệu văn bản", "upper", 13, 13, "upright", false, "II.3.c"),
  diaDanhNgayThang: element("diaDanhNgayThang", "Địa danh và thời gian ban hành", "lower", 13, 14, "italic", false, "II.4.c"),
  tenLoai: element("tenLoai", "Tên loại văn bản", "upper", 13, 14, "upright", true, "II.5.b"),
  trichYeu: element("trichYeu", "Trích yếu nội dung văn bản", "lower", 13, 14, "upright", true, "II.5.b"),
  trichYeuCongVan: element("trichYeuCongVan", "Trích yếu nội dung công văn (sau \"V/v\")", "lower", 12, 13, "upright", false, "II.5.b"),
  canCu: element("canCu", "Căn cứ ban hành văn bản", "lower", 13, 14, "italic", false, "II.6.a"),
  noiDung: element("noiDung", "Nội dung văn bản", "lower", 13, 14, "upright", false, "II.6.e"),
  phanChuong: element("phanChuong", "Từ \"Phần\", \"Chương\" và số thứ tự", "lower", 13, 14, "upright", true, "II.6.đ"),
  tieuDePhanChuong: element("tieuDePhanChuong", "Tiêu đề của phần, chương", "upper", 13, 14, "upright", true, "II.6.đ"),
  mucTieuMuc: element("mucTieuMuc", "Từ \"Mục\", \"Tiểu mục\" và số thứ tự", "lower", 13, 14, "upright", true, "II.6.đ"),
  tieuDeMucTieuMuc: element("tieuDeMucTieuMuc", "Tiêu đề của mục, tiểu mục", "upper", 13, 14, "upright", true, "II.6.đ"),
  dieu: element("dieu", "Từ \"Điều\", số thứ tự và tiêu đề của điều", "lower", 13, 14, "upright", true, "II.6.đ"),
  khoan: element("khoan", "Khoản (không có tiêu đề)", "lower", 13, 14, "upright", false, "II.6.đ"),
  khoanCoTieuDe: element("khoanCoTieuDe", "Số thứ tự và tiêu đề của khoản (khi có tiêu đề)", "lower", 13, 14, "upright", true, "II.6.đ"),
  diem: element("diem", "Điểm", "lower", 13, 14, "upright", false, "II.6.đ"),
  quyenHanChucVu: element("quyenHanChucVu", "Quyền hạn và chức vụ của người ký", "upper", 13, 14, "upright", true, "II.7.đ"),
  hoTenNguoiKy: element("hoTenNguoiKy", "Họ và tên của người ký", "lower", 13, 14, "upright", true, "II.7.đ"),
  kinhGui: element("kinhGui", "Từ \"Kính gửi\" và tên nơi nhận", "lower", 13, 14, "upright", false, "II.9.d"),
  noiNhanTieuDe: element("noiNhanTieuDe", "Từ \"Nơi nhận\"", "lower", 12, 12, "italic", true, "II.9.d"),
  noiNhanDanhSach: element("noiNhanDanhSach", "Danh sách nơi nhận", "lower", 11, 11, "upright", false, "II.9.d"),
  phuLucSo: element("phuLucSo", "Từ \"Phụ lục\" và số thứ tự", "lower", 14, 14, "upright", true, "III.1.b"),
  phuLucTieuDe: element("phuLucTieuDe", "Tên phụ lục", "upper", 13, 14, "upright", true, "III.1.b"),
  phuLucChiDan: element("phuLucChiDan", "Chỉ dẫn kèm theo văn bản trên phụ lục", "lower", 13, 14, "italic", false, "III.1.c"),
  dauKhan: element("dauKhan", "Dấu chỉ mức độ khẩn", "upper", 13, 14, "upright", true, "III.2.b"),
  phamViLuuHanh: element("phamViLuuHanh", "Chỉ dẫn về phạm vi lưu hành", "upper", 13, 14, "upright", true, "III.2.c"),
  kyHieuNguoiSoan: element("kyHieuNguoiSoan", "Ký hiệu người soạn thảo và số lượng bản phát hành", "upper", 11, 11, "upright", false, "III.3"),
  diaChiCoQuan: element("diaChiCoQuan", "Địa chỉ cơ quan, thư điện tử, trang thông tin điện tử, số điện thoại, số Fax", "lower", 11, 12, "upright", false, "III.4"),
};

export const PAGE = {
  paper: {
    name: "A4",
    widthMm: 210,
    heightMm: 297,
    source: `${PL1} I.1`,
  },
  /** Lề cách mép giấy, đơn vị mm. */
  marginsMm: {
    top: { min: 20, max: 25, source: `${PL1} I.3` },
    bottom: { min: 20, max: 25, source: `${PL1} I.3` },
    left: { min: 30, max: 35, source: `${PL1} I.3` },
    right: { min: 15, max: 20, source: `${PL1} I.3` },
  },
};

export type MarginSide = keyof typeof PAGE.marginsMm;

export const FONT = {
  name: "Times New Roman",
  encoding: "Unicode TCVN 6909:2001",
  color: "#000000",
  source: `${PL1} I.4`,
};

export const PAGE_NUMBER = {
  sizePt: { min: 13, max: 14 },
  align: "center" as const,
  position: "top-margin" as const,
  hideOnFirstPage: true,
  source: `${PL1} I.7`,
};

export const CONTENT_LAYOUT = {
  align: "justify" as const,
  /** Chữ đầu dòng lùi vào 1 cm hoặc 1,27 cm. */
  firstLineIndentMm: [10, 12.7],
  paragraphSpacingMinPt: 6,
  /** Khoảng cách giữa các dòng: tối thiểu dòng đơn, tối đa 1,5 lines. */
  lineSpacingLines: { min: 1, max: 1.5 },
  source: `${PL1} II.6.e`,
};

/** Độ dài đường kẻ ngang so với độ dài dòng chữ liền trên. */
export const LINE_RULES = {
  tieuNgu: { lengthRatio: { min: 1, max: 1 }, source: `${PL1} II.1.b` },
  tenCoQuanBanHanh: { lengthRatio: { min: 1 / 3, max: 1 / 2 }, source: `${PL1} II.2.b` },
  trichYeu: { lengthRatio: { min: 1 / 3, max: 1 / 2 }, source: `${PL1} II.5.b` },
};

export interface SignAuthority {
  abbr: string;
  meaning: string;
}

export const SIGN_AUTHORITY: readonly SignAuthority[] = [
  { abbr: "TM.", meaning: "ký thay mặt tập thể" },
  { abbr: "Q.", meaning: "quyền cấp trưởng" },
  { abbr: "KT.", meaning: "ký thay người đứng đầu" },
  { abbr: "TL.", meaning: "ký thừa lệnh" },
  { abbr: "TUQ.", meaning: "ký thừa uỷ quyền" },
];

export interface UrgencyMark {
  text: string;
  widthMm: number;
  heightMm: number;
}

/** Con dấu chỉ mức độ khẩn: khung chữ nhật viền đơn (Phụ lục I, Phần I, Mục III.2.b). */
export const URGENCY_MARKS: readonly UrgencyMark[] = [
  { text: "HỎA TỐC", widthMm: 30, heightMm: 8 },
  { text: "THƯỢNG KHẨN", widthMm: 40, heightMm: 8 },
  { text: "KHẨN", widthMm: 20, heightMm: 8 },
];
