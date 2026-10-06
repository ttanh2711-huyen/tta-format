import { readdirSync, readFileSync, statSync } from "node:fs";
import { join, relative, sep } from "node:path";
import { getStandard, STANDARDS } from "./index";
import { HD05 } from "./hd05/profile";
import { ND30 } from "./nd30/profile";

describe("khai báo hai chuẩn", () => {
  it("có đúng hai chuẩn, tra cứu được theo mã", () => {
    expect(STANDARDS.map((s) => s.id)).toEqual(["nd30", "hd05"]);
    expect(getStandard("nd30")).toBe(ND30);
    expect(getStandard("hd05")).toBe(HD05);
  });

  it("mã chuẩn không tồn tại là lỗi", () => {
    expect(() => getStandard("khong-co" as never)).toThrow(/khong-co/);
  });

  it("thông tin nhận diện và căn cứ pháp lý", () => {
    expect(ND30).toMatchObject({
      name: "Chính quyền (NĐ 30)",
      identifyingTitle: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
      effectiveDate: "2020-03-05",
    });
    expect(ND30.legalBasis).toMatch(/30\/2020\/NĐ-CP/);
    expect(HD05).toMatchObject({
      name: "Đảng (HD 05)",
      identifyingTitle: "ĐẢNG CỘNG SẢN VIỆT NAM",
      effectiveDate: "2026-05-27",
    });
    expect(HD05.legalBasis).toMatch(/05-HD\/VPTW/);
  });

  it("hồ sơ có đủ quy tắc, danh sách loại văn bản theo từng chuẩn", () => {
    expect(ND30.docTypes).toHaveLength(29);
    expect(HD05.docTypes).toHaveLength(33);
    expect(ND30.signAuthority.map((a) => a.abbr)).toContain("TUQ.");
    expect(HD05.signAuthority.map((a) => a.abbr)).not.toContain("TUQ.");
    expect(Object.keys(ND30.elements).length).toBeGreaterThan(0);
    expect(Object.keys(HD05.elements).length).toBeGreaterThan(0);
  });

  it("hai chuẩn khác nhau ở các điểm chính (đối chiếu mục 6B của kế hoạch)", () => {
    expect(ND30.elements.chucVu).toBeUndefined();
    expect(HD05.elements.chucVu?.bold).toBe(false);
    expect(ND30.elements.quyenHanChucVu?.bold).toBe(true);
    expect(ND30.content.lineSpacing.kind).toBe("multiple");
    expect(HD05.content.lineSpacing.kind).toBe("exactly");
    expect(ND30.page.duplexBackMarginsMm).toBeUndefined();
    expect(HD05.page.duplexBackMarginsMm).toBeDefined();
  });
});

describe("formatSoKyHieu theo từng chuẩn", () => {
  it("chuẩn Chính quyền: \"Số: 05/QĐ-UBND\", công văn không có chữ viết tắt tên loại", () => {
    expect(ND30.formatSoKyHieu(15, "QĐ", ["BNV"])).toBe("Số: 15/QĐ-BNV");
    expect(ND30.formatSoKyHieu(5, null, ["BNV", "VP"])).toBe("Số: 05/BNV-VP");
    expect(ND30.formatSoKyHieu(7, "CTr", ["UBND"])).toBe("Số: 07/CTr-UBND");
    expect(ND30.formatSoKyHieu(12, null, ["UBND", "VX"])).toBe("Số: 12/UBND-VX");
  });

  it("chuẩn Đảng: \"Số 05-HD/VPTW\", không có dấu hai chấm, số và ký hiệu nối bằng \"-\"", () => {
    expect(HD05.formatSoKyHieu(5, "HD", ["VPTW"])).toBe("Số 05-HD/VPTW");
    expect(HD05.formatSoKyHieu(127, "QĐ", ["TW"])).toBe("Số 127-QĐ/TW");
    expect(HD05.formatSoKyHieu(23, "QĐ", ["UBKTTW"])).toBe("Số 23-QĐ/UBKTTW");
    expect(HD05.formatSoKyHieu(138, "CV", ["BTGDVTU"])).toBe("Số 138-CV/BTGDVTU");
  });

  it("chuẩn Đảng: văn bản liên cơ quan nối các chữ viết tắt cơ quan bằng \"-\"", () => {
    expect(HD05.formatSoKyHieu(5, "QC", ["BTCTU", "BNCTU"])).toBe("Số 05-QC/BTCTU-BNCTU");
  });

  it("chuẩn Đảng: văn bản đại hội dùng ký hiệu chung ĐH", () => {
    expect(HD05.formatSoKyHieu(1, "BC", ["ĐH"])).toBe("Số 01-BC/ĐH");
    expect(HD05.formatSoKyHieu(20, "BB", ["ĐH"])).toBe("Số 20-BB/ĐH");
  });

  it("chuẩn Đảng bắt buộc có chữ viết tắt tên loại", () => {
    expect(() => HD05.formatSoKyHieu(5, null, ["TW"])).toThrow(RangeError);
  });

  it.each([ND30, HD05])("%# từ chối số không hợp lệ, thiếu cơ quan, chữ viết tắt có khoảng trắng hoặc ký tự lạ", (standard) => {
    expect(() => standard.formatSoKyHieu(0, "QĐ", ["UBND"])).toThrow(RangeError);
    expect(() => standard.formatSoKyHieu(5, "QĐ", [])).toThrow(RangeError);
    expect(() => standard.formatSoKyHieu(5, "QĐ", ["UB ND"])).toThrow(RangeError);
    expect(() => standard.formatSoKyHieu(5, "Q/Đ", ["UBND"])).toThrow(RangeError);
    expect(() => standard.formatSoKyHieu(5, "QĐ", [""])).toThrow(RangeError);
  });
});

describe("mã dùng chung không rẽ nhánh theo mã chuẩn", () => {
  const srcRoot = join(__dirname, "..");
  const branching = /(===|!==|==|!=)\s*["'](nd30|hd05)["']|["'](nd30|hd05)["']\s*(===|!==|==|!=)|case\s+["'](nd30|hd05)["']/;

  function listSourceFiles(dir: string): string[] {
    return readdirSync(dir).flatMap((name) => {
      const full = join(dir, name);
      if (statSync(full).isDirectory()) {
        return listSourceFiles(full);
      }
      return full.endsWith(".ts") && !full.endsWith(".test.ts") ? [full] : [];
    });
  }

  it("không có `=== \"nd30\"`, `case \"hd05\"`... trong mã nguồn ngoài thư mục riêng của từng chuẩn", () => {
    const offenders = listSourceFiles(srcRoot)
      .filter((file) => {
        const rel = relative(srcRoot, file).split(sep).join("/");
        return !rel.startsWith("standards/nd30/") && !rel.startsWith("standards/hd05/");
      })
      .filter((file) => branching.test(readFileSync(file, "utf8")))
      .map((file) => relative(srcRoot, file));
    expect(offenders).toEqual([]);
  });
});
