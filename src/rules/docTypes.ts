/**
 * Các hình thức văn bản hành chính (Điều 7 Nghị định 30/2020/NĐ-CP), chữ viết tắt tên
 * loại (Phụ lục III, Mục I) và mẫu trình bày (Phụ lục III, Mục II).
 *
 * Danh sách 29 loại lấy từ Điều 7 qua nguồn thứ cấp, đối chiếu với 27 loại có chữ viết
 * tắt ở Phụ lục III (đọc nguyên văn). Công văn và thư công không có chữ viết tắt tên loại.
 */

export type LayoutFamily =
  | "nghi-quyet" // Mẫu 1.1
  | "quyet-dinh" // Mẫu 1.2, 1.3
  | "co-ten-loai" // Mẫu 1.4: văn bản có tên loại và trích yếu
  | "cong-van" // Mẫu 1.5
  | "cong-dien" // Mẫu 1.6
  | "giay-moi" // Mẫu 1.7
  | "giay-gioi-thieu" // Mẫu 1.8
  | "bien-ban" // Mẫu 1.9
  | "giay-nghi-phep" // Mẫu 1.10
  | "chua-co-mau"; // Phụ lục III không có mẫu riêng

export interface DocType {
  id: string;
  name: string;
  /** Chữ viết tắt tên loại dùng trong ký hiệu văn bản; null nếu không có. */
  abbr: string | null;
  layout: LayoutFamily;
  /** Số mẫu ở Phụ lục III, Mục II; null nếu không có mẫu. */
  template: string | null;
}

function docType(
  id: string,
  name: string,
  abbr: string | null,
  layout: LayoutFamily,
  template: string | null,
): DocType {
  return { id, name, abbr, layout, template };
}

const MAU_1_4 = "Mẫu 1.4";

export const DOC_TYPES: readonly DocType[] = [
  docType("nghi-quyet", "Nghị quyết (cá biệt)", "NQ", "nghi-quyet", "Mẫu 1.1"),
  docType("quyet-dinh", "Quyết định (cá biệt)", "QĐ", "quyet-dinh", "Mẫu 1.2, Mẫu 1.3"),
  docType("chi-thi", "Chỉ thị", "CT", "co-ten-loai", MAU_1_4),
  docType("quy-che", "Quy chế", "QC", "co-ten-loai", MAU_1_4),
  docType("quy-dinh", "Quy định", "QyĐ", "co-ten-loai", MAU_1_4),
  docType("thong-cao", "Thông cáo", "TC", "co-ten-loai", MAU_1_4),
  docType("thong-bao", "Thông báo", "TB", "co-ten-loai", MAU_1_4),
  docType("huong-dan", "Hướng dẫn", "HD", "co-ten-loai", MAU_1_4),
  docType("chuong-trinh", "Chương trình", "CTr", "co-ten-loai", MAU_1_4),
  docType("ke-hoach", "Kế hoạch", "KH", "co-ten-loai", MAU_1_4),
  docType("phuong-an", "Phương án", "PA", "co-ten-loai", MAU_1_4),
  docType("de-an", "Đề án", "ĐA", "co-ten-loai", MAU_1_4),
  docType("du-an", "Dự án", "DA", "co-ten-loai", MAU_1_4),
  docType("bao-cao", "Báo cáo", "BC", "co-ten-loai", MAU_1_4),
  docType("bien-ban", "Biên bản", "BB", "bien-ban", "Mẫu 1.9"),
  docType("to-trinh", "Tờ trình", "TTr", "co-ten-loai", MAU_1_4),
  docType("hop-dong", "Hợp đồng", "HĐ", "chua-co-mau", null),
  docType("cong-van", "Công văn", null, "cong-van", "Mẫu 1.5"),
  docType("cong-dien", "Công điện", "CĐ", "cong-dien", "Mẫu 1.6"),
  docType("ban-ghi-nho", "Bản ghi nhớ", "BGN", "chua-co-mau", null),
  docType("ban-thoa-thuan", "Bản thỏa thuận", "BTT", "chua-co-mau", null),
  docType("giay-uy-quyen", "Giấy ủy quyền", "GUQ", "co-ten-loai", MAU_1_4),
  docType("giay-moi", "Giấy mời", "GM", "giay-moi", "Mẫu 1.7"),
  docType("giay-gioi-thieu", "Giấy giới thiệu", "GGT", "giay-gioi-thieu", "Mẫu 1.8"),
  docType("giay-nghi-phep", "Giấy nghỉ phép", "GNP", "giay-nghi-phep", "Mẫu 1.10"),
  docType("phieu-gui", "Phiếu gửi", "PG", "co-ten-loai", MAU_1_4),
  docType("phieu-chuyen", "Phiếu chuyển", "PC", "co-ten-loai", MAU_1_4),
  docType("phieu-bao", "Phiếu báo", "PB", "co-ten-loai", MAU_1_4),
  docType("thu-cong", "Thư công", null, "chua-co-mau", null),
];

export interface CopyType {
  name: string;
  abbr: string;
}

/** Bản sao văn bản (Phụ lục III, Mục I). Ngoài phạm vi add-in, giữ lại để tra cứu. */
export const COPY_TYPES: readonly CopyType[] = [
  { name: "Bản sao y", abbr: "SY" },
  { name: "Bản trích sao", abbr: "TrS" },
  { name: "Bản sao lục", abbr: "SL" },
];

export function findDocTypeById(id: string): DocType | undefined {
  return DOC_TYPES.find((t) => t.id === id);
}

export function findDocTypeByAbbr(abbr: string): DocType | undefined {
  return DOC_TYPES.find((t) => t.abbr === abbr);
}
