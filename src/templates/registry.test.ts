import { STANDARDS } from "../standards";
import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { findTemplate, templatesFor } from "./registry";

describe("registry loại văn bản", () => {
  it.each(STANDARDS.map((profile) => [profile.name, profile] as const))("%s: mọi loại đều có mục hợp lệ hoặc cờ chưa hỗ trợ kèm lý do", (_name, profile) => {
    const entries = templatesFor(profile);
    for (const type of profile.docTypes) expect(entries.some((entry) => entry.docTypeId === type.id)).toBe(true);
    for (const entry of entries) {
      expect(entry.key).not.toBe("");
      if (entry.supported) expect(entry.family).not.toBeNull();
      else expect(entry.reason ?? "").not.toBe("");
    }
    expect(new Set(entries.map((entry) => entry.key)).size).toBe(entries.length);
  });

  it("chuẩn Chính quyền: 17 loại dùng Mẫu 1.4, trong đó quy chế và quy định có cấu trúc Chương/Điều", () => {
    const titled = templatesFor(ND30).filter((entry) => entry.family === "titled");
    expect(titled).toHaveLength(15);
    expect(templatesFor(ND30).filter((entry) => entry.family === "regulation").map((entry) => entry.docTypeId)).toEqual(["quy-che", "quy-dinh"]);
    expect(titled.length + 2).toBe(ND30.docTypes.filter((type) => type.layout === "co-ten-loai").length);
    expect(titled.map((entry) => entry.docTypeId)).toEqual(expect.arrayContaining(["chi-thi", "thong-bao", "ke-hoach", "bao-cao", "to-trinh", "giay-uy-quyen", "phieu-gui", "phieu-chuyen", "phieu-bao"]));
  });

  it("chữ viết tắt gợi ý theo từng loại", () => {
    const abbr = (key: string) => findTemplate(ND30, key)?.abbr;
    expect(abbr("thong-bao")).toBe("TB");
    expect(abbr("to-trinh")).toBe("TTr");
    expect(abbr("quy-dinh")).toBe("QyĐ");
    expect(abbr("cong-van")).toBeNull();
    expect(findTemplate(HD05, "to-trinh")?.abbr).toBe("TTr");
  });

  it("công văn dùng họ công văn; quyết định cá biệt có hai biến thể, nghị quyết một", () => {
    expect(findTemplate(ND30, "cong-van")?.family).toBe("congVan");
    expect(findTemplate(ND30, "quyet-dinh:quyetDinhTrucTiep")?.decisionKind).toBe("quyetDinhTrucTiep");
    expect(findTemplate(ND30, "quyet-dinh:quyetDinhGianTiep")?.decisionKind).toBe("quyetDinhGianTiep");
    expect(findTemplate(ND30, "nghi-quyet")?.decisionKind).toBe("nghiQuyet");
  });

  it("chuẩn Đảng: tờ trình có Kính trình, công văn riêng, thể loại còn lại có tên loại", () => {
    expect(findTemplate(HD05, "to-trinh")?.family).toBe("toTrinh");
    expect(findTemplate(HD05, "cong-van")?.family).toBe("congVan");
    expect(findTemplate(HD05, "bao-cao")?.family).toBe("titled");
  });

  it("biên bản và các văn bản nhiều bên ký có họ mẫu riêng", () => {
    expect(findTemplate(ND30, "bien-ban")?.family).toBe("minutes");
    for (const id of ["hop-dong", "ban-ghi-nho", "ban-thoa-thuan"]) expect(findTemplate(ND30, id)?.family).toBe("agreement");
  });

  it("loại chưa có họ mẫu hiện kèm lý do, không bị ẩn", () => {
    const entry = findTemplate(ND30, "cong-dien");
    expect(entry).toMatchObject({ supported: false, family: null });
    expect(entry?.reason).toContain("chưa có mẫu");
  });
});
