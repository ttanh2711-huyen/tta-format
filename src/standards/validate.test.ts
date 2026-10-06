import { ptToMm } from "./units";
import { HD05 } from "./hd05/profile";
import { ND30 } from "./nd30/profile";
import {
  isFirstLineIndentValid,
  isFontSizeValid,
  isLineSpacingValid,
  isMarginValid,
} from "./validate";

describe("isMarginValid: chuẩn Chính quyền (khoảng giá trị)", () => {
  it("chấp nhận lề nằm trong khoảng, kể cả hai đầu mút", () => {
    expect(isMarginValid(ND30, "top", 20)).toBe(true);
    expect(isMarginValid(ND30, "top", 25)).toBe(true);
    expect(isMarginValid(ND30, "bottom", 22.5)).toBe(true);
    expect(isMarginValid(ND30, "left", 30)).toBe(true);
    expect(isMarginValid(ND30, "left", 35)).toBe(true);
    expect(isMarginValid(ND30, "right", 15)).toBe(true);
    expect(isMarginValid(ND30, "right", 20)).toBe(true);
  });

  it("từ chối lề ngoài khoảng", () => {
    expect(isMarginValid(ND30, "top", 19.9)).toBe(false);
    expect(isMarginValid(ND30, "top", 25.1)).toBe(false);
    expect(isMarginValid(ND30, "left", 29.9)).toBe(false);
    expect(isMarginValid(ND30, "left", 35.1)).toBe(false);
    expect(isMarginValid(ND30, "right", 14.9)).toBe(false);
    expect(isMarginValid(ND30, "right", 20.1)).toBe(false);
  });

  it("chịu được sai số làm tròn khi đổi từ pt sang mm", () => {
    expect(isMarginValid(ND30, "left", ptToMm(85.04))).toBe(true);
    expect(isMarginValid(ND30, "right", ptToMm(42.5))).toBe(true);
  });

  it("từ chối giá trị không phải số hữu hạn", () => {
    expect(isMarginValid(ND30, "top", Number.NaN)).toBe(false);
    expect(isMarginValid(ND30, "top", Number.POSITIVE_INFINITY)).toBe(false);
  });

  it("chuẩn không quy định lề mặt sau thì mặt sau dùng chung lề mặt trước", () => {
    expect(isMarginValid(ND30, "left", 15, { duplexBack: true })).toBe(false);
    expect(isMarginValid(ND30, "left", 30, { duplexBack: true })).toBe(true);
  });
});

describe("isMarginValid: chuẩn Đảng (giá trị cố định, in hai mặt)", () => {
  it("mặt trước: trên 20, dưới 20, trái 30, phải 15 mm", () => {
    expect(isMarginValid(HD05, "top", 20)).toBe(true);
    expect(isMarginValid(HD05, "bottom", 20)).toBe(true);
    expect(isMarginValid(HD05, "left", 30)).toBe(true);
    expect(isMarginValid(HD05, "right", 15)).toBe(true);
  });

  it("từ chối giá trị khác, kể cả giá trị nằm trong khoảng của chuẩn Chính quyền", () => {
    expect(isMarginValid(HD05, "top", 25)).toBe(false);
    expect(isMarginValid(HD05, "left", 35)).toBe(false);
    expect(isMarginValid(HD05, "right", 20)).toBe(false);
  });

  it("lề trái 15 mm không hợp lệ ở mặt trước nhưng hợp lệ ở mặt sau khi in hai mặt", () => {
    expect(isMarginValid(HD05, "left", 15)).toBe(false);
    expect(isMarginValid(HD05, "left", 15, { duplexBack: true })).toBe(true);
    expect(isMarginValid(HD05, "right", 30, { duplexBack: true })).toBe(true);
  });

  it("mặt sau: lề trái 30 mm và phải 15 mm không còn hợp lệ", () => {
    expect(isMarginValid(HD05, "left", 30, { duplexBack: true })).toBe(false);
    expect(isMarginValid(HD05, "right", 15, { duplexBack: true })).toBe(false);
  });
});

describe("isFontSizeValid", () => {
  it("chuẩn Chính quyền: quốc hiệu cỡ 12-13", () => {
    expect(isFontSizeValid(ND30, "quocHieu", 12)).toBe(true);
    expect(isFontSizeValid(ND30, "quocHieu", 13)).toBe(true);
    expect(isFontSizeValid(ND30, "quocHieu", 11)).toBe(false);
    expect(isFontSizeValid(ND30, "quocHieu", 14)).toBe(false);
  });

  it("chuẩn Đảng: tiêu đề cỡ 15, thể loại cỡ 15-16", () => {
    expect(isFontSizeValid(HD05, "tieuDe", 15)).toBe(true);
    expect(isFontSizeValid(HD05, "tieuDe", 14)).toBe(false);
    expect(isFontSizeValid(HD05, "theLoai", 15)).toBe(true);
    expect(isFontSizeValid(HD05, "theLoai", 16)).toBe(true);
    expect(isFontSizeValid(HD05, "theLoai", 17)).toBe(false);
  });

  it("cho phép cỡ chữ lẻ nửa điểm nằm trong khoảng, từ chối giá trị không hữu hạn", () => {
    expect(isFontSizeValid(ND30, "noiDung", 13.5)).toBe(true);
    expect(isFontSizeValid(ND30, "noiDung", Number.NaN)).toBe(false);
  });

  it("thành phần không có trong chuẩn là lỗi lập trình, không âm thầm trả về false", () => {
    expect(() => isFontSizeValid(ND30, "tieuDe", 15)).toThrow(/tieuDe/);
    expect(() => isFontSizeValid(HD05, "quocHieu", 13)).toThrow(/quocHieu/);
  });
});

describe("isLineSpacingValid", () => {
  it("chuẩn Đảng: Exactly 18-22 pt", () => {
    expect(isLineSpacingValid(HD05, { kind: "exactly", pt: 17 })).toBe(false);
    expect(isLineSpacingValid(HD05, { kind: "exactly", pt: 18 })).toBe(true);
    expect(isLineSpacingValid(HD05, { kind: "exactly", pt: 20 })).toBe(true);
    expect(isLineSpacingValid(HD05, { kind: "exactly", pt: 22 })).toBe(true);
    expect(isLineSpacingValid(HD05, { kind: "exactly", pt: 22.5 })).toBe(false);
  });

  it("chuẩn Đảng: giãn dòng theo bội số (Single, 1,5 lines) không hợp lệ", () => {
    expect(isLineSpacingValid(HD05, { kind: "multiple", lines: 1 })).toBe(false);
    expect(isLineSpacingValid(HD05, { kind: "multiple", lines: 1.5 })).toBe(false);
  });

  it("chuẩn Chính quyền: từ dòng đơn đến 1,5 lines, ví dụ 1,17 của file mẫu", () => {
    expect(isLineSpacingValid(ND30, { kind: "multiple", lines: 0.9 })).toBe(false);
    expect(isLineSpacingValid(ND30, { kind: "multiple", lines: 1 })).toBe(true);
    expect(isLineSpacingValid(ND30, { kind: "multiple", lines: 1.17 })).toBe(true);
    expect(isLineSpacingValid(ND30, { kind: "multiple", lines: 1.5 })).toBe(true);
    expect(isLineSpacingValid(ND30, { kind: "multiple", lines: 1.6 })).toBe(false);
  });

  it("chuẩn Chính quyền: giãn dòng Exactly không có trong quy định nên không hợp lệ", () => {
    expect(isLineSpacingValid(ND30, { kind: "exactly", pt: 20 })).toBe(false);
  });
});

describe("isFirstLineIndentValid", () => {
  it("chuẩn Chính quyền: 1 cm hoặc 1,27 cm", () => {
    expect(isFirstLineIndentValid(ND30, 10)).toBe(true);
    expect(isFirstLineIndentValid(ND30, 12.7)).toBe(true);
    expect(isFirstLineIndentValid(ND30, 12.65)).toBe(true);
    expect(isFirstLineIndentValid(ND30, 12)).toBe(false);
    expect(isFirstLineIndentValid(ND30, 0)).toBe(false);
  });

  it("chuẩn Đảng: khoảng 10 mm, cho sai số 1 mm", () => {
    expect(isFirstLineIndentValid(HD05, 10)).toBe(true);
    expect(isFirstLineIndentValid(HD05, 10.9)).toBe(true);
    expect(isFirstLineIndentValid(HD05, 12.7)).toBe(false);
  });
});
