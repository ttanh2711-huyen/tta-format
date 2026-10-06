import {
  CONTENT_LAYOUT,
  ELEMENTS,
  FONT,
  LINE_RULES,
  PAGE,
  PAGE_NUMBER,
  SIGN_AUTHORITY,
  URGENCY_MARKS,
  type ElementId,
} from "./rules";
import type { Casing, FontStyle } from "../types";

describe("NĐ 30: khổ giấy và lề trang (Phụ lục I, Phần I, Mục I)", () => {
  it("khổ giấy là A4 210 x 297 mm", () => {
    expect(PAGE.paper).toMatchObject({ name: "A4", widthMm: 210, heightMm: 297 });
  });

  it("lề trên và dưới 20-25 mm, trái 30-35 mm, phải 15-20 mm", () => {
    expect(PAGE.marginsMm.top).toMatchObject({ min: 20, max: 25 });
    expect(PAGE.marginsMm.bottom).toMatchObject({ min: 20, max: 25 });
    expect(PAGE.marginsMm.left).toMatchObject({ min: 30, max: 35 });
    expect(PAGE.marginsMm.right).toMatchObject({ min: 15, max: 20 });
  });

  it("phông chữ Times New Roman, Unicode TCVN 6909:2001, màu đen", () => {
    expect(FONT).toMatchObject({
      name: "Times New Roman",
      encoding: "Unicode TCVN 6909:2001",
      color: "#000000",
    });
  });

  it("số trang đặt giữa lề trên, cỡ 13-14, không hiển thị ở trang đầu", () => {
    expect(PAGE_NUMBER).toMatchObject({
      sizePt: { min: 13, max: 14 },
      align: "center",
      hideOnFirstPage: true,
    });
  });
});

describe("NĐ 30: cỡ chữ và kiểu chữ từng thành phần (Phụ lục I, Phần I, Mục II-V)", () => {
  const rows: Array<[ElementId, Casing, number, number, FontStyle, boolean]> = [
    ["quocHieu", "upper", 12, 13, "upright", true],
    ["tieuNgu", "lower", 13, 14, "upright", true],
    ["tenCoQuanChuQuan", "upper", 12, 13, "upright", false],
    ["tenCoQuanBanHanh", "upper", 12, 13, "upright", true],
    ["so", "lower", 13, 13, "upright", false],
    ["kyHieu", "upper", 13, 13, "upright", false],
    ["diaDanhNgayThang", "lower", 13, 14, "italic", false],
    ["tenLoai", "upper", 13, 14, "upright", true],
    ["trichYeu", "lower", 13, 14, "upright", true],
    ["trichYeuCongVan", "lower", 12, 13, "upright", false],
    ["canCu", "lower", 13, 14, "italic", false],
    ["noiDung", "lower", 13, 14, "upright", false],
    ["phanChuong", "lower", 13, 14, "upright", true],
    ["tieuDePhanChuong", "upper", 13, 14, "upright", true],
    ["mucTieuMuc", "lower", 13, 14, "upright", true],
    ["tieuDeMucTieuMuc", "upper", 13, 14, "upright", true],
    ["dieu", "lower", 13, 14, "upright", true],
    ["khoan", "lower", 13, 14, "upright", false],
    ["khoanCoTieuDe", "lower", 13, 14, "upright", true],
    ["diem", "lower", 13, 14, "upright", false],
    ["quyenHanChucVu", "upper", 13, 14, "upright", true],
    ["hoTenNguoiKy", "lower", 13, 14, "upright", true],
    ["kinhGui", "lower", 13, 14, "upright", false],
    ["noiNhanTieuDe", "lower", 12, 12, "italic", true],
    ["noiNhanDanhSach", "lower", 11, 11, "upright", false],
    ["phuLucSo", "lower", 14, 14, "upright", true],
    ["phuLucTieuDe", "upper", 13, 14, "upright", true],
    ["phuLucChiDan", "lower", 13, 14, "italic", false],
    ["dauKhan", "upper", 13, 14, "upright", true],
    ["phamViLuuHanh", "upper", 13, 14, "upright", true],
    ["kyHieuNguoiSoan", "upper", 11, 11, "upright", false],
    ["diaChiCoQuan", "lower", 11, 12, "upright", false],
  ];

  it.each(rows)("%s: %s, cỡ %i-%i, %s, đậm=%s", (id, casing, min, max, style, bold) => {
    const rule = ELEMENTS[id];
    expect(rule.casing).toBe(casing);
    expect(rule.sizePt).toEqual({ min, max });
    expect(rule.style).toBe(style);
    expect(rule.bold).toBe(bold);
  });

  it("khai báo đủ các thành phần trong bảng trên và không thừa thành phần nào", () => {
    expect(Object.keys(ELEMENTS).sort()).toEqual(rows.map((r) => r[0]).sort());
  });

  it("mọi thành phần có nhãn tiếng Việt, cỡ chữ hợp lệ và trích dẫn Phụ lục I", () => {
    for (const [id, rule] of Object.entries(ELEMENTS)) {
      expect(rule.label.trim()).not.toBe("");
      expect(rule.sizePt.min).toBeLessThanOrEqual(rule.sizePt.max);
      expect(rule.source).toMatch(/^Phụ lục I,/);
      expect(id).toBe(rule.id);
    }
  });
});

describe("NĐ 30: phần nội dung, đường kẻ, quyền hạn ký, dấu khẩn", () => {
  it("nội dung canh đều hai lề, thụt dòng đầu 1 cm hoặc 1,27 cm, giãn đoạn tối thiểu 6 pt, giãn dòng từ đơn đến 1,5", () => {
    expect(CONTENT_LAYOUT).toMatchObject({
      align: "justify",
      firstLineIndentMm: [10, 12.7],
      paragraphSpacingMinPt: 6,
      lineSpacing: { kind: "multiple", min: 1, max: 1.5 },
    });
  });

  it("đường kẻ dưới tiêu ngữ bằng độ dài dòng chữ; dưới tên cơ quan và trích yếu dài 1/3 đến 1/2 dòng chữ", () => {
    expect(LINE_RULES.tieuNgu).toMatchObject({ lengthRatio: { min: 1, max: 1 } });
    expect(LINE_RULES.tenCoQuanBanHanh.lengthRatio.min).toBeCloseTo(1 / 3, 5);
    expect(LINE_RULES.tenCoQuanBanHanh.lengthRatio.max).toBe(0.5);
    expect(LINE_RULES.trichYeu.lengthRatio.min).toBeCloseTo(1 / 3, 5);
    expect(LINE_RULES.trichYeu.lengthRatio.max).toBe(0.5);
  });

  it("quyền hạn ký gồm TM., Q., KT., TL., TUQ.", () => {
    expect(SIGN_AUTHORITY.map((a) => a.abbr)).toEqual(["TM.", "Q.", "KT.", "TL.", "TUQ."]);
  });

  it("dấu mức độ khẩn: HỎA TỐC 30x8 mm, THƯỢNG KHẨN 40x8 mm, KHẨN 20x8 mm", () => {
    expect(URGENCY_MARKS).toEqual([
      expect.objectContaining({ text: "HỎA TỐC", widthMm: 30, heightMm: 8 }),
      expect.objectContaining({ text: "THƯỢNG KHẨN", widthMm: 40, heightMm: 8 }),
      expect.objectContaining({ text: "KHẨN", widthMm: 20, heightMm: 8 }),
    ]);
  });
});
