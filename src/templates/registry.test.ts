import { STANDARDS } from "../standards";
import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { findTemplate, templatesFor, visibleTemplates } from "./registry";
import { findSmallForm } from "./smallForms";

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
    for (const id of ["nghi-quyet", "quyet-dinh", "quy-dinh", "quy-che"]) expect(findTemplate(HD05, id)?.family).toBe("regulation");
  });

  it("biên bản và các văn bản nhiều bên ký có họ mẫu riêng", () => {
    expect(findTemplate(ND30, "bien-ban")?.family).toBe("minutes");
    for (const id of ["hop-dong", "ban-ghi-nho", "ban-thoa-thuan"]) expect(findTemplate(ND30, id)?.family).toBe("agreement");
  });

  it("giấy tờ nhỏ của chuẩn Chính quyền có họ mẫu form; mọi loại của chuẩn này đều được hỗ trợ", () => {
    for (const id of ["giay-moi", "giay-gioi-thieu", "cong-dien", "giay-nghi-phep", "thu-cong"]) expect(findTemplate(ND30, id)?.family).toBe("form");
    expect(templatesFor(ND30).filter((entry) => !entry.supported)).toEqual([]);
  });

  it("chuẩn Đảng: biên bản, giấy tờ, văn bản đại hội, liên cơ quan đều có họ mẫu", () => {
    expect(findTemplate(HD05, "bien-ban")?.family).toBe("minutes");
    for (const id of ["giay-chung-nhan", "giay-di-duong", "giay-nghi-phep", "giay-moi", "phieu-chuyen", "phieu-gui", "giay-gioi-thieu", "thu-cong"]) expect(findTemplate(HD05, id)?.family).toBe("form");
    expect(findTemplate(HD05, "van-ban-dai-hoi")).toMatchObject({ family: "congress", abbr: "ĐH" });
    expect(findTemplate(HD05, "lien-co-quan")?.family).toBe("joint");
    expect(templatesFor(HD05).filter((entry) => !entry.supported)).toEqual([]);
  });

  it("giấy tờ ít dùng của chuẩn Đảng bị ẩn khỏi ô chọn nhưng vẫn có cấu hình; chuẩn Chính quyền không ẩn gì", () => {
    const hidden = ["giay-chung-nhan", "giay-di-duong", "phieu-chuyen", "phieu-gui"];
    const visible = visibleTemplates(HD05).map((entry) => entry.docTypeId);
    for (const id of hidden) {
      expect(findTemplate(HD05, id)).toMatchObject({ family: "form", supported: true, hidden: true });
      expect(visible).not.toContain(id);
    }
    expect(visible).toEqual(expect.arrayContaining(["giay-moi", "giay-gioi-thieu", "giay-nghi-phep", "thu-cong", "cong-van", "bao-cao"]));
    expect(visibleTemplates(ND30)).toEqual(templatesFor(ND30));
    expect(visibleTemplates(ND30).map((entry) => entry.docTypeId)).toEqual(expect.arrayContaining(["phieu-chuyen", "phieu-gui"]));
  });

  it("mọi giấy tờ có họ mẫu form đều có cấu hình mục cần điền", () => {
    for (const profile of [ND30, HD05]) for (const entry of templatesFor(profile).filter((e) => e.family === "form")) expect(findSmallForm(entry.docTypeId)).toBeDefined();
  });

  it("chuẩn Chính quyền có bản sao y, trích sao, sao lục (ký hiệu SY, TrS, SL) dùng họ mẫu form; chuẩn Đảng không có", () => {
    for (const [key, abbr] of [["sao-y", "SY"], ["trich-sao", "TrS"], ["sao-luc", "SL"]] as const) {
      expect(findTemplate(ND30, key)).toMatchObject({ family: "form", supported: true, abbr });
      expect(findSmallForm(key)).toBeDefined();
      expect(findTemplate(HD05, key)).toBeUndefined();
    }
  });

  it("loại chưa có họ mẫu hiện kèm lý do, không bị ẩn (cơ chế)", () => {
    const entry = findTemplate({ ...ND30, docTypes: [...ND30.docTypes, { id: "x", name: "X", abbr: null, layout: "la", template: null }] } as typeof ND30, "x");
    expect(entry).toMatchObject({ supported: false, family: null });
    expect(entry?.reason).toContain("chưa có mẫu");
  });
});
