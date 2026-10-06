import type { Casing, FontStyle } from "../types";
import {
  CONTENT_LAYOUT,
  ELEMENTS,
  FONT,
  HEADER_MARKS,
  PAGE,
  PAGE_NUMBER,
  SIGN_AUTHORITY,
  URGENCY_MARKS,
  type ElementId,
} from "./rules";

describe("HD 05: khổ giấy và lề (Phần II, Mục IV)", () => {
  it("khổ A4 210 x 297 mm, trình bày theo chiều dọc", () => {
    expect(PAGE.paper).toMatchObject({ name: "A4", widthMm: 210, heightMm: 297 });
  });

  it("lề cố định: trên 20, dưới 20, trái 30, phải 15 mm", () => {
    expect(PAGE.marginsMm.top).toMatchObject({ min: 20, max: 20 });
    expect(PAGE.marginsMm.bottom).toMatchObject({ min: 20, max: 20 });
    expect(PAGE.marginsMm.left).toMatchObject({ min: 30, max: 30 });
    expect(PAGE.marginsMm.right).toMatchObject({ min: 15, max: 15 });
  });

  it("in hai mặt: mặt sau trên 20, dưới 20, trái 15, phải 30 mm", () => {
    expect(PAGE.duplexBackMarginsMm?.top).toMatchObject({ min: 20, max: 20 });
    expect(PAGE.duplexBackMarginsMm?.bottom).toMatchObject({ min: 20, max: 20 });
    expect(PAGE.duplexBackMarginsMm?.left).toMatchObject({ min: 15, max: 15 });
    expect(PAGE.duplexBackMarginsMm?.right).toMatchObject({ min: 30, max: 30 });
  });

  it("cho phép trình bày theo chiều ngang khi có bảng, biểu", () => {
    expect(PAGE.landscapeAllowedForTables).toBe(true);
  });

  it("phông Times New Roman, Unicode TCVN 6909:2001, màu đen", () => {
    expect(FONT).toMatchObject({
      name: "Times New Roman",
      encoding: "Unicode TCVN 6909:2001",
      color: "#000000",
    });
  });

  it("số trang: giữa, cách mép trên 10 mm, cỡ 14, ẩn ở trang đầu, phụ lục đánh số riêng", () => {
    expect(PAGE_NUMBER).toMatchObject({
      sizePt: { min: 14, max: 14 },
      align: "center",
      position: "top-margin",
      distanceFromTopEdgeMm: 10,
      hideOnFirstPage: true,
      appendixNumberedSeparately: true,
    });
  });
});

describe("HD 05: phần nội dung (Phần II, Mục I.6.2.2)", () => {
  it("canh đều hai lề, thụt dòng đầu khoảng 10 mm, giãn đoạn tối thiểu 6 pt, giãn dòng 18-22 pt kiểu Exactly, kết thúc bằng dấu chấm", () => {
    expect(CONTENT_LAYOUT).toMatchObject({
      align: "justify",
      firstLineIndentMm: [10],
      paragraphSpacingMinPt: 6,
      lineSpacing: { kind: "exactly", minPt: 18, maxPt: 22 },
      endsWithPeriod: true,
    });
    expect(CONTENT_LAYOUT.firstLineIndentToleranceMm).toBeGreaterThan(0);
  });
});

describe("HD 05: cỡ chữ và kiểu chữ từng thành phần (Phụ lục 2)", () => {
  const rows: Array<[ElementId, Casing, number, number, FontStyle, boolean, boolean]> = [
    // id, loại chữ, cỡ nhỏ nhất, cỡ lớn nhất, kiểu chữ, đậm, gạch chân
    ["tieuDe", "upper", 15, 15, "upright", true, false],
    ["coQuanCapTren", "upper", 14, 14, "upright", false, false],
    ["coQuanBanHanh", "upper", 14, 14, "upright", true, false],
    ["soKyHieu", "lower", 14, 14, "upright", false, false],
    ["diaDanhNgayThang", "lower", 14, 14, "italic", false, false],
    ["theLoai", "upper", 15, 16, "upright", true, false],
    ["trichYeu", "lower", 14, 15, "upright", true, false],
    ["trichYeuCongVan", "lower", 12, 12, "italic", false, false],
    ["noiDung", "lower", 14, 15, "upright", false, false],
    ["phanChuong", "lower", 14, 15, "upright", true, false],
    ["tenPhanChuong", "upper", 14, 15, "upright", true, false],
    ["mucSo", "lower", 14, 15, "upright", true, false],
    ["tenMuc", "upper", 14, 15, "upright", true, false],
    ["dieu", "lower", 14, 15, "upright", true, false],
    ["khoan", "lower", 14, 15, "upright", false, false],
    ["diem", "lower", 14, 15, "upright", false, false],
    ["quyenHan", "upper", 14, 14, "upright", true, false],
    ["chucVu", "upper", 14, 14, "upright", false, false],
    ["hoTen", "lower", 14, 14, "upright", true, false],
    ["kinhGuiTu", "lower", 14, 14, "italic", false, false],
    ["kinhGuiDanhSach", "lower", 14, 15, "upright", false, false],
    ["noiNhanTu", "lower", 14, 14, "upright", false, true],
    ["noiNhanDanhSach", "lower", 12, 12, "upright", false, false],
    ["dauKhan", "upper", 14, 14, "upright", true, false],
    ["phamViLuuHanh", "upper", 12, 12, "upright", true, false],
    ["duThao", "lower", 14, 14, "italic", false, false],
    ["kyHieuNguoiSoan", "lower", 11, 11, "upright", false, false],
    ["thongTinLienHe", "lower", 10, 10, "upright", false, false],
    ["soTrang", "lower", 14, 14, "upright", false, false],
  ];

  it.each(rows)("%s: %s, cỡ %i-%i, %s, đậm=%s, gạch chân=%s", (id, casing, min, max, style, bold, underline) => {
    const rule = ELEMENTS[id];
    expect(rule.casing).toBe(casing);
    expect(rule.sizePt).toEqual({ min, max });
    expect(rule.style).toBe(style);
    expect(rule.bold).toBe(bold);
    expect(rule.underline === true).toBe(underline);
  });

  it("khai báo đủ các thành phần trong bảng và không thừa", () => {
    expect(Object.keys(ELEMENTS).sort()).toEqual(rows.map((r) => r[0]).sort());
  });

  it("mọi thành phần có nhãn tiếng Việt, cỡ chữ hợp lệ và trích dẫn Hướng dẫn 05", () => {
    for (const [id, rule] of Object.entries(ELEMENTS)) {
      expect(rule.label.trim()).not.toBe("");
      expect(rule.sizePt.min).toBeLessThanOrEqual(rule.sizePt.max);
      expect(rule.source).toMatch(/^HD 05, /);
      expect(id).toBe(rule.id);
    }
  });

  it("chức vụ người ký không đậm, khác với chuẩn Chính quyền", () => {
    expect(ELEMENTS.chucVu.bold).toBe(false);
    expect(ELEMENTS.quyenHan.bold).toBe(true);
  });
});

describe("HD 05: dấu hiệu nhận biết và quy cách riêng của chuẩn Đảng", () => {
  it("tiêu đề, dấu sao dưới tên cơ quan, 5 dấu gạch nối dưới trích yếu", () => {
    expect(HEADER_MARKS).toMatchObject({
      identifyingTitle: "ĐẢNG CỘNG SẢN VIỆT NAM",
      agencyStar: "*",
      trichYeuDashCount: 5,
    });
  });

  it("quyền hạn ký: T/M, K/T, T/L, Q.", () => {
    expect(SIGN_AUTHORITY.map((a) => a.abbr)).toEqual(["T/M", "K/T", "T/L", "Q."]);
  });

  it("mức độ khẩn: khẩn, thượng khẩn, hỏa tốc", () => {
    expect(URGENCY_MARKS.map((m) => m.text)).toEqual(["KHẨN", "THƯỢNG KHẨN", "HỎA TỐC"]);
  });
});
