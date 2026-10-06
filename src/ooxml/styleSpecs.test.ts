import { isFontSizeValid } from "../standards/validate";
import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { mmToPt } from "../standards/units";
import { buildStyleSpecs, defaultStyleOptions, type StyleSpec } from "./styleSpecs";

function byName(specs: StyleSpec[], name: string): StyleSpec {
  const found = specs.find((s) => s.name === name);
  if (!found) {
    throw new Error(`Không có style ${name}; có: ${specs.map((s) => s.name).join(", ")}`);
  }
  return found;
}

describe("buildStyleSpecs: tên style", () => {
  it("chuẩn Chính quyền: tiền tố ND30_, đủ các style nội dung và bảng", () => {
    const names = buildStyleSpecs(ND30).map((s) => s.name);
    expect(names).toEqual(
      expect.arrayContaining([
        "ND30_NoiDung",
        "ND30_KinhGui",
        "ND30_CanCu",
        "ND30_PhanChuong",
        "ND30_TieuDePhanChuong",
        "ND30_MucTieuMuc",
        "ND30_TieuDeMucTieuMuc",
        "ND30_Dieu",
        "ND30_Khoan",
        "ND30_KhoanCoTieuDe",
        "ND30_Diem",
        "ND30_PhuLucSo",
        "ND30_PhuLucTieuDe",
        "ND30_PhuLucChiDan",
        "ND30_BangDuLieu",
      ]),
    );
  });

  it("chuẩn Đảng: tiền tố HD05_, có các cấp Phần/Chương/Mục/Điều/Khoản/Điểm", () => {
    const names = buildStyleSpecs(HD05).map((s) => s.name);
    expect(names).toEqual(
      expect.arrayContaining([
        "HD05_NoiDung",
        "HD05_KinhGuiTu",
        "HD05_KinhGuiDanhSach",
        "HD05_PhanChuong",
        "HD05_TenPhanChuong",
        "HD05_MucSo",
        "HD05_TenMuc",
        "HD05_Dieu",
        "HD05_Khoan",
        "HD05_Diem",
        "HD05_BangDuLieu",
      ]),
    );
  });

  it.each([ND30, HD05])("%# tên không trùng nhau, đúng tiền tố, không đè style có sẵn của Word", (profile) => {
    const specs = buildStyleSpecs(profile);
    const names = specs.map((s) => s.name);
    expect(new Set(names).size).toBe(names.length);
    for (const name of names) {
      expect(name.startsWith(profile.styles.prefix)).toBe(true);
      expect(name).not.toMatch(/^(Normal|Heading|Title|Body|Tiêu đề|Bình thường)/i);
    }
  });

  it("hai chuẩn không dùng chung tên style nên cùng tồn tại trong một tài liệu", () => {
    const nd30 = new Set(buildStyleSpecs(ND30).map((s) => s.name));
    for (const spec of buildStyleSpecs(HD05)) {
      expect(nd30.has(spec.name)).toBe(false);
    }
  });
});

describe("buildStyleSpecs: thuộc tính chuẩn Chính quyền (theo file mẫu)", () => {
  const specs = buildStyleSpecs(ND30);

  it("ND30_NoiDung: Times New Roman 14, canh đều, thụt 1,27 cm, giãn 6/6 pt, giãn dòng 1,17", () => {
    const s = byName(specs, "ND30_NoiDung");
    expect(s.font).toMatchObject({ name: "Times New Roman", sizePt: 14, bold: false, italic: false, color: "#000000" });
    expect(s.paragraph).toMatchObject({
      alignment: "Justified",
      spaceBeforePt: 6,
      spaceAfterPt: 6,
      lineSpacing: { kind: "multiple", lines: 1.17 },
    });
    expect(s.paragraph.firstLineIndentPt).toBeCloseTo(mmToPt(12.7), 2);
  });

  it("ND30_KinhGui: canh giữa, không thụt, giãn 12/12 pt như file mẫu", () => {
    const s = byName(specs, "ND30_KinhGui");
    expect(s.paragraph).toMatchObject({ alignment: "Centered", firstLineIndentPt: 0, spaceBeforePt: 12, spaceAfterPt: 12 });
  });

  it("ND30_Dieu đậm; ND30_CanCu nghiêng; ND30_TieuDePhanChuong đậm canh giữa", () => {
    expect(byName(specs, "ND30_Dieu").font.bold).toBe(true);
    expect(byName(specs, "ND30_CanCu").font.italic).toBe(true);
    const heading = byName(specs, "ND30_TieuDePhanChuong");
    expect(heading.font.bold).toBe(true);
    expect(heading.paragraph.alignment).toBe("Centered");
  });

  it("ND30_BangDuLieu: canh trái, 14 pt, giãn 3/3 pt, giãn dòng đơn, không thụt", () => {
    const s = byName(specs, "ND30_BangDuLieu");
    expect(s.font.sizePt).toBe(14);
    expect(s.paragraph).toMatchObject({
      alignment: "Left",
      firstLineIndentPt: 0,
      spaceBeforePt: 3,
      spaceAfterPt: 3,
      lineSpacing: { kind: "multiple", lines: 1 },
    });
  });
});

describe("buildStyleSpecs: thuộc tính chuẩn Đảng", () => {
  const specs = buildStyleSpecs(HD05);

  it("HD05_NoiDung: canh đều, thụt 10 mm, giãn 6/6 pt, giãn dòng Exactly 20 pt", () => {
    const s = byName(specs, "HD05_NoiDung");
    expect(s.font).toMatchObject({ name: "Times New Roman", sizePt: 14 });
    expect(s.paragraph).toMatchObject({ alignment: "Justified", lineSpacing: { kind: "exactly", pt: 20 } });
    expect(s.paragraph.firstLineIndentPt).toBeCloseTo(mmToPt(10), 2);
  });

  it("mọi style nội dung của chuẩn Đảng dùng giãn dòng Exactly; bảng dùng giãn dòng đơn", () => {
    for (const s of specs.filter((x) => x.name !== "HD05_BangDuLieu")) {
      expect(s.paragraph.lineSpacing.kind).toBe("exactly");
    }
    expect(byName(specs, "HD05_BangDuLieu").paragraph.lineSpacing).toEqual({ kind: "multiple", lines: 1 });
  });

  it("HD05_KinhGuiTu nghiêng canh giữa; HD05_Dieu đậm", () => {
    expect(byName(specs, "HD05_KinhGuiTu").font.italic).toBe(true);
    expect(byName(specs, "HD05_Dieu").font.bold).toBe(true);
  });
});

describe("buildStyleSpecs: cỡ chữ luôn hợp lệ với chuẩn", () => {
  it.each([ND30, HD05])("%# cỡ chữ mặc định nằm trong khoảng của từng thành phần", (profile) => {
    for (const spec of buildStyleSpecs(profile)) {
      if (spec.elementId in profile.elements) {
        expect(isFontSizeValid(profile, spec.elementId, spec.font.sizePt)).toBe(true);
      }
    }
  });

  it("đổi cỡ chữ sang 13: style nội dung của chuẩn Chính quyền theo, chuẩn Đảng bị kéo về cỡ nhỏ nhất cho phép (14)", () => {
    expect(byName(buildStyleSpecs(ND30, { ...defaultStyleOptions(ND30), fontSizePt: 13 }), "ND30_NoiDung").font.sizePt).toBe(13);
    expect(byName(buildStyleSpecs(HD05, { ...defaultStyleOptions(HD05), fontSizePt: 13 }), "HD05_NoiDung").font.sizePt).toBe(14);
  });

  it("đổi giãn đoạn và thụt dòng chỉ ảnh hưởng style nội dung, không đổi giãn 12/12 của Kính gửi hay 3/3 của bảng", () => {
    const specs = buildStyleSpecs(ND30, { ...defaultStyleOptions(ND30), spaceBeforePt: 8, spaceAfterPt: 8, firstLineIndentMm: 10 });
    expect(byName(specs, "ND30_NoiDung").paragraph).toMatchObject({ spaceBeforePt: 8, spaceAfterPt: 8 });
    expect(byName(specs, "ND30_NoiDung").paragraph.firstLineIndentPt).toBeCloseTo(mmToPt(10), 2);
    expect(byName(specs, "ND30_KinhGui").paragraph).toMatchObject({ spaceBeforePt: 12, spaceAfterPt: 12 });
    expect(byName(specs, "ND30_BangDuLieu").paragraph).toMatchObject({ spaceBeforePt: 3, spaceAfterPt: 3 });
  });
});
