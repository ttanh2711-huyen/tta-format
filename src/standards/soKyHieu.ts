const ABBREVIATION = /^[\p{L}\p{N}]+$/u;

/**
 * Chữ viết tắt trong ký hiệu văn bản chỉ gồm chữ và số, không có khoảng trắng hay dấu
 * phân cách ("/" và "-" do chuẩn ghép vào). Hai chuẩn cùng quy định như vậy.
 */
export function assertAbbreviation(value: string, what: string): void {
  if (!ABBREVIATION.test(value)) {
    throw new RangeError(`${what} chỉ gồm chữ và số, không có khoảng trắng hoặc dấu phân cách: "${value}".`);
  }
}

/** Kiểm tra các phần tử cơ quan, đơn vị: phải có ít nhất một phần tử hợp lệ. */
export function assertAgencyAbbreviations(agencyAbbrs: readonly string[]): void {
  if (agencyAbbrs.length === 0) {
    throw new RangeError("Ký hiệu văn bản cần ít nhất một chữ viết tắt tên cơ quan, tổ chức.");
  }
  for (const abbr of agencyAbbrs) {
    assertAbbreviation(abbr, "Chữ viết tắt tên cơ quan, tổ chức");
  }
}
