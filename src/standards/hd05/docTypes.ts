/**
 * Thể loại văn bản của Đảng (Điều 5) và loại văn bản, giấy tờ hành chính (Điều 6) theo
 * Quy định 399-QĐ/TW ngày 09/01/2026 của Ban Bí thư; chữ viết tắt theo Hướng dẫn 05-HD/VPTW,
 * Phần II, Mục I.3.1.2.a.
 *
 * - `abbrSource: "explicit"`: Hướng dẫn 05 nêu rõ (kể cả qua ví dụ số, ký hiệu).
 * - `abbrSource: "derived"`: suy ra theo quy tắc "chữ cái đầu các âm tiết"; người dùng nên kiểm tra.
 *
 * Danh sách tên lấy từ hai nguồn thứ cấp trùng nhau (xem docs/hd05-ambiguities.md); chưa đối
 * chiếu toàn văn Quy định 399-QĐ/TW. Hướng dẫn 05 chỉ có một sơ đồ bố trí chung (Phụ lục 1),
 * không có mẫu riêng từng thể loại, nên `template` luôn là null.
 */

import type { DocType as GenericDocType } from "../types";

export type LayoutFamily = "cong-van" | "co-ten-loai" | "giay-to";

export type DocType = GenericDocType<LayoutFamily>;

function explicit(id: string, name: string, abbr: string, layout: LayoutFamily = "co-ten-loai"): DocType {
  return { id, name, abbr, layout, template: null, abbrSource: "explicit" };
}

function derived(id: string, name: string, abbr: string, layout: LayoutFamily = "co-ten-loai"): DocType {
  return { id, name, abbr, layout, template: null, abbrSource: "derived" };
}

function noAbbr(id: string, name: string, layout: LayoutFamily = "co-ten-loai", note?: string): DocType {
  return { id, name, abbr: null, layout, template: null, ...(note ? { note } : {}) };
}

export const DOC_TYPES: readonly DocType[] = [
  // Điều 5: thể loại văn bản
  noAbbr("cuong-linh-chinh-tri", "Cương lĩnh chính trị"),
  noAbbr("dieu-le-dang", "Điều lệ Đảng"),
  noAbbr("chien-luoc", "Chiến lược"),
  explicit("nghi-quyet", "Nghị quyết", "NQ"),
  explicit("quyet-dinh", "Quyết định", "QĐ"),
  explicit("chi-thi", "Chỉ thị", "CT"),
  explicit("ket-luan", "Kết luận", "KL"),
  explicit("quy-che", "Quy chế", "QC"),
  explicit("quy-dinh", "Quy định", "QĐ"),
  explicit("thong-tri", "Thông tri", "TT"),
  explicit("huong-dan", "Hướng dẫn", "HD"),
  derived("thong-bao", "Thông báo", "TB"),
  derived("thong-cao", "Thông cáo", "TC"),
  noAbbr("tuyen-bo", "Tuyên bố", "co-ten-loai", "Chữ cái đầu là TB, trùng với thông báo; Hướng dẫn 05 chưa quy định riêng."),
  noAbbr("loi-keu-goi", "Lời kêu gọi"),
  explicit("bao-cao", "Báo cáo", "BC"),
  derived("ke-hoach", "Kế hoạch", "KH"),
  derived("quy-hoach", "Quy hoạch", "QH"),
  explicit("chuong-trinh", "Chương trình", "CTr"),
  derived("de-an", "Đề án", "ĐA"),
  derived("phuong-an", "Phương án", "PA"),
  derived("du-an", "Dự án", "DA"),
  explicit("to-trinh", "Tờ trình", "TTr"),
  explicit("cong-van", "Công văn", "CV", "cong-van"),
  explicit("bien-ban", "Biên bản", "BB"),
  // Điều 6: văn bản, giấy tờ hành chính
  derived("giay-gioi-thieu", "Giấy giới thiệu", "GGT", "giay-to"),
  derived("giay-chung-nhan", "Giấy chứng nhận", "GCN", "giay-to"),
  derived("giay-di-duong", "Giấy đi đường", "GĐĐ", "giay-to"),
  derived("giay-nghi-phep", "Giấy nghỉ phép", "GNP", "giay-to"),
  derived("giay-moi", "Giấy mời", "GM", "giay-to"),
  derived("phieu-chuyen", "Phiếu chuyển", "PC", "giay-to"),
  derived("phieu-gui", "Phiếu gửi", "PG", "giay-to"),
  noAbbr("thu-cong", "Thư công", "giay-to"),
];

export function findDocTypeById(id: string): DocType | undefined {
  return DOC_TYPES.find((t) => t.id === id);
}

/** Chữ viết tắt không duy nhất (quyết định và quy định cùng là QĐ) nên trả về danh sách. */
export function findDocTypesByAbbr(abbr: string): DocType[] {
  return DOC_TYPES.filter((t) => t.abbr === abbr);
}
