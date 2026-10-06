/**
 * Bảng chữ viết tắt tên cơ quan ban hành văn bản của Đảng dùng trong ký hiệu văn bản
 * (HD 05, Phần II, Mục I.3.1.2.b.iv; đối chiếu bằng ảnh trang 8-9 của văn bản gốc).
 * Tên cấp uỷ cụ thể (TU, TW, VPTU, BTCTU...) là chữ viết tắt do cấp uỷ tự đặt theo quy tắc
 * chữ cái đầu, không nằm trong bảng này.
 */

export type AgencyGroup = "dang-uy-chi-bo" | "co-quan-chuyen-trach" | "ban-chi-dao" | "dai-hoi" | "don-vi-su-nghiep";

export interface AgencyCode {
  name: string;
  code: string;
  group: AgencyGroup;
}

function agency(name: string, code: string, group: AgencyGroup): AgencyCode {
  return { name, code, group };
}

export const AGENCY_CODES: readonly AgencyCode[] = [
  agency("Đảng uỷ quân sự", "ĐUQS", "dang-uy-chi-bo"),
  agency("Quân uỷ Trung ương", "QUTW", "dang-uy-chi-bo"),
  agency("Đảng uỷ công an", "ĐUCA", "dang-uy-chi-bo"),
  agency("Đảng uỷ biên phòng", "ĐUBP", "dang-uy-chi-bo"),
  agency("Các đảng uỷ khác", "ĐU", "dang-uy-chi-bo"),
  agency("Chi bộ", "CB", "dang-uy-chi-bo"),

  agency("Uỷ ban Kiểm tra", "UBKT", "co-quan-chuyen-trach"),
  agency("Cơ quan uỷ ban kiểm tra", "CQUBKT", "co-quan-chuyen-trach"),
  agency("Ban tuyên giáo và dân vận", "BTGDV", "co-quan-chuyen-trach"),
  agency("Ban xây dựng đảng", "BXDĐ", "co-quan-chuyen-trach"),

  agency("Ban chỉ đạo", "BCĐ", "ban-chi-dao"),
  agency("Tiểu ban", "TB", "ban-chi-dao"),
  agency("Hội đồng", "HĐ", "ban-chi-dao"),

  agency("Đại hội", "ĐH", "dai-hoi"),

  agency("Báo Nhân Dân", "BND", "don-vi-su-nghiep"),
  agency("Tạp chí Cộng sản", "TCCS", "don-vi-su-nghiep"),
  agency("Học viện Chính trị quốc gia Hồ Chí Minh", "HVCTQGHCM", "don-vi-su-nghiep"),
  agency("Đài Truyền hình Việt Nam", "ĐTHVN", "don-vi-su-nghiep"),
  agency("Đài Tiếng nói Việt Nam", "ĐTNVN", "don-vi-su-nghiep"),
  agency("Thông tấn xã Việt Nam", "TTXVN", "don-vi-su-nghiep"),
  agency("Viện Hàn lâm Khoa học xã hội Việt Nam", "VHLKHXHVN", "don-vi-su-nghiep"),
  agency("Viện Hàn lâm Khoa học và Công nghệ Việt Nam", "VHLKHCNVN", "don-vi-su-nghiep"),
  agency("Trường chính trị", "TCT", "don-vi-su-nghiep"),
  agency("Trung tâm chính trị", "TTCT", "don-vi-su-nghiep"),
  agency("Cơ quan báo và phát thanh truyền hình", "CQBPTTH", "don-vi-su-nghiep"),
];

/** Tên trong văn bản gốc viết "uỷ"; chấp nhận cả cách bỏ dấu "ủy" khi tra cứu. */
function normalize(text: string): string {
  return text.normalize("NFC").toLowerCase().replace(/ủy/g, "uỷ");
}

export function findAgencyByCode(code: string): AgencyCode | undefined {
  return AGENCY_CODES.find((a) => a.code === code);
}

export function findAgencyByName(name: string): AgencyCode | undefined {
  const wanted = normalize(name);
  return AGENCY_CODES.find((a) => normalize(a.name) === wanted);
}
