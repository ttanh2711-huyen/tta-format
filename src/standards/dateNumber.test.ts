import { formatNgay, formatSo } from "./dateNumber";

describe("formatSo", () => {
  it("số nhỏ hơn 10 ghi thêm số 0 phía trước", () => {
    expect(formatSo(1)).toBe("01");
    expect(formatSo(7)).toBe("07");
    expect(formatSo(9)).toBe("09");
  });

  it("số từ 10 trở lên giữ nguyên", () => {
    expect(formatSo(10)).toBe("10");
    expect(formatSo(123)).toBe("123");
  });

  it("từ chối số không phải số nguyên dương", () => {
    expect(() => formatSo(0)).toThrow(RangeError);
    expect(() => formatSo(-3)).toThrow(RangeError);
    expect(() => formatSo(2.5)).toThrow(RangeError);
    expect(() => formatSo(Number.NaN)).toThrow(RangeError);
  });
});

describe("formatNgay", () => {
  it("ngày nhỏ hơn 10 và tháng 1, 2 ghi thêm số 0", () => {
    expect(formatNgay(5, 1, 2020)).toBe("ngày 05 tháng 01 năm 2020");
    expect(formatNgay(9, 2, 2026)).toBe("ngày 09 tháng 02 năm 2026");
  });

  it("tháng từ 3 đến 9 không thêm số 0", () => {
    expect(formatNgay(29, 6, 2019)).toBe("ngày 29 tháng 6 năm 2019");
    expect(formatNgay(10, 3, 2026)).toBe("ngày 10 tháng 3 năm 2026");
  });

  it("tháng 10 đến 12 giữ nguyên, ngày nhỏ hơn 10 vẫn thêm số 0", () => {
    expect(formatNgay(3, 10, 2026)).toBe("ngày 03 tháng 10 năm 2026");
    expect(formatNgay(31, 12, 2026)).toBe("ngày 31 tháng 12 năm 2026");
  });

  it("chấp nhận 29/2 của năm nhuận, ví dụ ở HD 05: ngày 03 tháng 02 năm 2026", () => {
    expect(formatNgay(29, 2, 2024)).toBe("ngày 29 tháng 02 năm 2024");
    expect(formatNgay(3, 2, 2026)).toBe("ngày 03 tháng 02 năm 2026");
  });

  it("từ chối ngày không tồn tại", () => {
    expect(() => formatNgay(31, 2, 2026)).toThrow(RangeError);
    expect(() => formatNgay(29, 2, 2025)).toThrow(RangeError);
    expect(() => formatNgay(31, 4, 2026)).toThrow(RangeError);
    expect(() => formatNgay(0, 1, 2026)).toThrow(RangeError);
  });

  it("từ chối tháng hoặc năm không hợp lệ", () => {
    expect(() => formatNgay(1, 13, 2026)).toThrow(RangeError);
    expect(() => formatNgay(1, 0, 2026)).toThrow(RangeError);
    expect(() => formatNgay(1, 1, 999)).toThrow(RangeError);
    expect(() => formatNgay(1, 1, 2026.5)).toThrow(RangeError);
  });
});
