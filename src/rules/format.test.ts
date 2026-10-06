import { formatNgay, formatSo, isMarginValid, mmToPt, ptToMm } from "./format";

describe("đổi đơn vị mm và pt", () => {
  it("25,4 mm bằng 72 pt (1 inch)", () => {
    expect(mmToPt(25.4)).toBeCloseTo(72, 6);
    expect(ptToMm(72)).toBeCloseTo(25.4, 6);
  });

  it("lề trái 30 mm bằng khoảng 85,04 pt", () => {
    expect(mmToPt(30)).toBeCloseTo(85.0394, 3);
  });

  it("đổi qua lại không làm lệch giá trị", () => {
    expect(ptToMm(mmToPt(20))).toBeCloseTo(20, 9);
  });
});

describe("isMarginValid (Phụ lục I, Phần I, Mục I.3)", () => {
  it("chấp nhận lề nằm trong khoảng, kể cả hai đầu mút", () => {
    expect(isMarginValid("top", 20)).toBe(true);
    expect(isMarginValid("top", 25)).toBe(true);
    expect(isMarginValid("bottom", 22.5)).toBe(true);
    expect(isMarginValid("left", 30)).toBe(true);
    expect(isMarginValid("left", 35)).toBe(true);
    expect(isMarginValid("right", 15)).toBe(true);
    expect(isMarginValid("right", 20)).toBe(true);
  });

  it("từ chối lề ngoài khoảng", () => {
    expect(isMarginValid("top", 19.9)).toBe(false);
    expect(isMarginValid("top", 25.1)).toBe(false);
    expect(isMarginValid("left", 29.9)).toBe(false);
    expect(isMarginValid("left", 35.1)).toBe(false);
    expect(isMarginValid("right", 14.9)).toBe(false);
    expect(isMarginValid("right", 20.1)).toBe(false);
  });

  it("chịu được sai số làm tròn khi đổi từ pt sang mm", () => {
    expect(isMarginValid("left", ptToMm(85.04))).toBe(true);
    expect(isMarginValid("right", ptToMm(42.5))).toBe(true);
  });

  it("từ chối giá trị không phải số hữu hạn", () => {
    expect(isMarginValid("top", Number.NaN)).toBe(false);
    expect(isMarginValid("top", Number.POSITIVE_INFINITY)).toBe(false);
  });
});

describe("formatSo (Phụ lục I, Phần I, Mục II.3.c)", () => {
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

describe("formatNgay (Phụ lục I, Phần I, Mục II.4.b)", () => {
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

  it("chấp nhận 29/2 của năm nhuận", () => {
    expect(formatNgay(29, 2, 2024)).toBe("ngày 29 tháng 02 năm 2024");
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
