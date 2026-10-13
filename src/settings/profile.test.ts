import { STANDARDS } from "../standards";
import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { listAgencies, saveAgency, saveRecipients } from "./agencies";
import { saveDefaultStandard } from "./defaultStandard";
import { CONFIG_FORMAT, CONFIG_VERSION, exportConfig, importConfig } from "./exportImport";
import { buildRegistry, initialState } from "./formatOptions";
import { savePreset } from "./presets";
import { agencyLimits, defaultProfile, deleteProfile, listProfiles, saveProfile, setDefaultProfile, validateProfile, type ProfileInput } from "./profile";
import { memoryStore } from "./storage";

const nd30Input: ProfileInput = {
  standardId: "nd30",
  label: "Chi cục Thủy lợi",
  parentLines: ["Sở Nông nghiệp và Môi trường", "Thành phố Cần Thơ"],
  issuerLines: ["Chi cục Thủy lợi"],
  place: "Cần Thơ",
  abbrs: "CCTL-ĐĐPCTT",
  positionLines: ["Chi cục trưởng"],
  signerName: "Mai Hiếu Hiền",
};

const partyInput: ProfileInput = { ...nd30Input, standardId: "hd05", label: "Ban Tổ chức", parentLines: ["Tỉnh ủy Đồng Tháp"], issuerLines: ["Ban Tổ chức"], abbrs: "BTCTU" };

describe("kiểm tra hồ sơ theo giới hạn của chuẩn", () => {
  it("hồ sơ đúng thì không lỗi", () => {
    expect(validateProfile(ND30, nd30Input)).toEqual([]);
    expect(validateProfile(HD05, partyInput)).toEqual([]);
  });

  it("giới hạn dòng khác nhau giữa hai chuẩn", () => {
    expect(agencyLimits(ND30)).toEqual({ parent: 3, issuer: 1 });
    expect(agencyLimits(HD05)).toEqual({ parent: 2, issuer: 3 });
  });

  it.each([
    ["thiếu tên gợi nhớ", ND30, { label: " " }],
    ["thiếu cơ quan ban hành", ND30, { issuerLines: [""] }],
    ["Chính quyền: hai dòng cơ quan ban hành", ND30, { issuerLines: ["A", "B"] }],
    ["Chính quyền: bốn dòng cơ quan chủ quản", ND30, { parentLines: ["a", "b", "c", "d"] }],
    ["Đảng: ba dòng cấp trên", HD05, { parentLines: ["a", "b", "c"] }],
    ["Đảng: bốn dòng cơ quan ban hành", HD05, { issuerLines: ["a", "b", "c", "d"] }],
    ["chức vụ quá 3 dòng", ND30, { positionLines: ["a", "b", "c", "d"] }],
    ["viết tắt có khoảng trắng", ND30, { abbrs: "CC TL" }],
    ["viết tắt có dấu gạch chéo", ND30, { abbrs: "CCTL/ĐĐ" }],
  ])("%s bị từ chối kèm giải thích", (_name, standard, patch) => {
    const reasons = validateProfile(standard, { ...nd30Input, ...patch });
    expect(reasons.length).toBeGreaterThan(0);
    expect(reasons.every((reason) => /[a-zà-ỹ]/i.test(reason))).toBe(true);
  });

  it("viết tắt nối bằng gạch nối thì hợp lệ", () => {
    expect(validateProfile(ND30, { ...nd30Input, abbrs: "CCTL-ĐĐPCTT" })).toEqual([]);
  });
});

describe("lưu, sửa, xoá, mặc định", () => {
  it("lưu rồi đọc lại đúng dữ liệu, theo từng chuẩn", () => {
    const store = memoryStore();
    const saved = saveProfile(store, ND30, nd30Input);
    expect(saved.ok).toBe(true);
    saveProfile(store, HD05, partyInput);
    expect(listProfiles(store, ND30).map((p) => p.label)).toEqual(["Chi cục Thủy lợi"]);
    expect(listProfiles(store, HD05).map((p) => p.label)).toEqual(["Ban Tổ chức"]);
    expect(listProfiles(store, ND30)[0]).toMatchObject({ parentLines: nd30Input.parentLines, issuerLines: nd30Input.issuerLines, abbrs: "CCTL-ĐĐPCTT", signerName: "Mai Hiếu Hiền" });
  });

  it("giá trị ngoài giới hạn không được lưu", () => {
    const store = memoryStore();
    const result = saveProfile(store, ND30, { ...nd30Input, parentLines: ["a", "b", "c", "d"] });
    expect(result.ok).toBe(false);
    expect(listProfiles(store, ND30)).toEqual([]);
  });

  it("không trùng tên trong cùng chuẩn; trùng tên ở chuẩn khác thì được", () => {
    const store = memoryStore();
    saveProfile(store, ND30, nd30Input);
    expect(saveProfile(store, ND30, { ...nd30Input, label: "chi cục thủy lợi" }).ok).toBe(false);
    expect(saveProfile(store, HD05, { ...partyInput, label: "Chi cục Thủy lợi" }).ok).toBe(true);
  });

  it("sửa theo id giữ nguyên id và cờ mặc định", () => {
    const store = memoryStore();
    const first = saveProfile(store, ND30, nd30Input);
    if (!first.ok) throw new Error("lưu lỗi");
    setDefaultProfile(store, first.profile.id);
    const edited = saveProfile(store, ND30, { ...nd30Input, id: first.profile.id, signerName: "Người khác" });
    expect(edited.ok).toBe(true);
    expect(listProfiles(store, ND30)).toHaveLength(1);
    expect(listProfiles(store, ND30)[0]).toMatchObject({ id: first.profile.id, signerName: "Người khác", isDefault: true });
    expect(saveProfile(store, ND30, { ...nd30Input, id: "khong-co" }).ok).toBe(false);
  });

  it("đặt mặc định chỉ một hồ sơ mỗi chuẩn, không ảnh hưởng chuẩn khác", () => {
    const store = memoryStore();
    const a = saveProfile(store, ND30, nd30Input);
    const b = saveProfile(store, ND30, { ...nd30Input, label: "Văn phòng" });
    const c = saveProfile(store, HD05, partyInput);
    if (!a.ok || !b.ok || !c.ok) throw new Error("lưu lỗi");
    setDefaultProfile(store, a.profile.id);
    setDefaultProfile(store, c.profile.id);
    setDefaultProfile(store, b.profile.id);
    expect(defaultProfile(store, ND30)?.id).toBe(b.profile.id);
    expect(listProfiles(store, ND30).filter((p) => p.isDefault)).toHaveLength(1);
    expect(defaultProfile(store, HD05)?.id).toBe(c.profile.id);
    expect(setDefaultProfile(store, "khong-co")).toBe(false);
  });

  it("xoá hồ sơ; dữ liệu hỏng trong kho không làm vỡ danh sách", () => {
    const store = memoryStore();
    const a = saveProfile(store, ND30, nd30Input);
    if (!a.ok) throw new Error("lưu lỗi");
    deleteProfile(store, a.profile.id);
    expect(listProfiles(store, ND30)).toEqual([]);
    store.set("nd30.agencyProfiles", "{không phải json");
    expect(listProfiles(store, ND30)).toEqual([]);
  });
});

describe("xuất và nhập cấu hình", () => {
  function filledStore() {
    const store = memoryStore();
    const a = saveProfile(store, ND30, nd30Input);
    if (!a.ok) throw new Error("lưu lỗi");
    setDefaultProfile(store, a.profile.id);
    saveProfile(store, HD05, partyInput);
    saveAgency(store, ND30, "Sở Nội vụ");
    saveAgency(store, ND30, "CCTL", "abbrs");
    saveDefaultStandard(store, "hd05");
    const registry = buildRegistry(ND30);
    savePreset(store, ND30, "Của tôi", initialState(registry));
    return store;
  }

  it("xuất rồi nhập sang kho trống giữ nguyên hồ sơ, preset, danh sách cơ quan, chuẩn mặc định", () => {
    const text = exportConfig(filledStore(), new Date("2026-10-09T00:00:00Z"));
    const parsed = JSON.parse(text) as { format: string; version: number; exportedAt: string };
    expect(parsed).toMatchObject({ format: CONFIG_FORMAT, version: CONFIG_VERSION, exportedAt: "2026-10-09T00:00:00.000Z" });

    const target = memoryStore();
    const report = importConfig(target, text);
    expect(report).toMatchObject({ ok: true, profiles: 2, presets: 1, agencyEntries: 2 });
    expect(report.skipped).toEqual([]);
    expect(listProfiles(target, ND30)[0]).toMatchObject({ label: "Chi cục Thủy lợi", isDefault: true });
    expect(listProfiles(target, HD05)).toHaveLength(1);
    expect(exportConfig(target, new Date("2026-10-09T00:00:00Z"))).toBe(text);
  });

  it("danh bạ nơi nhận được xuất và nhập cùng cấu hình", () => {
    const source = memoryStore();
    saveRecipients(source, ND30, ["Sở Nội vụ", "Sở Tài chính"]);
    const target = memoryStore();
    saveRecipients(target, ND30, ["Sở Tài chính", "Sở Y tế"]);
    const report = importConfig(target, exportConfig(source));
    expect(report.ok).toBe(true);
    expect(listAgencies(target, ND30, "recipients")).toEqual(["Sở Tài chính", "Sở Y tế", "Sở Nội vụ"]);
  });

  it("nhập hai lần không nhân đôi", () => {
    const text = exportConfig(filledStore());
    const target = memoryStore();
    importConfig(target, text);
    importConfig(target, text);
    expect(listProfiles(target, ND30)).toHaveLength(1);
    expect(JSON.parse(exportConfig(target)).agencies.nd30.names).toEqual(["Sở Nội vụ"]);
  });

  it("hồ sơ trong file được thay theo id; hồ sơ sẵn có khác giữ nguyên", () => {
    const source = filledStore();
    const text = exportConfig(source);
    const target = memoryStore();
    saveProfile(target, ND30, { ...nd30Input, label: "Có sẵn trên máy này" });
    importConfig(target, text);
    expect(listProfiles(target, ND30).map((p) => p.label).sort()).toEqual(["Chi cục Thủy lợi", "Có sẵn trên máy này"]);
  });

  it("file hỏng, sai loại hoặc phiên bản mới hơn bị từ chối và không ghi gì", () => {
    const target = memoryStore();
    expect(importConfig(target, "không phải json").ok).toBe(false);
    expect(importConfig(target, JSON.stringify({ hello: 1 })).error).toContain("không phải file cấu hình");
    expect(importConfig(target, JSON.stringify({ format: CONFIG_FORMAT })).error).toContain("phiên bản");
    const newer = importConfig(target, JSON.stringify({ format: CONFIG_FORMAT, version: CONFIG_VERSION + 1, profiles: [] }));
    expect(newer.ok).toBe(false);
    expect(newer.error).toContain("cập nhật add-in");
    expect(listProfiles(target, ND30)).toEqual([]);
  });

  it("hồ sơ sai cấu trúc, sai chuẩn hoặc ngoài giới hạn bị bỏ qua và báo lại", () => {
    const target = memoryStore();
    const bad = { ...nd30Input, id: "x1", parentLines: ["a", "b", "c", "d"] };
    const unknown = { ...nd30Input, id: "x2", standardId: "khong-co" };
    const file = { format: CONFIG_FORMAT, version: 1, profiles: [bad, unknown, { id: "x3" }, { ...nd30Input, id: "ok1" }], presets: [{ nonsense: true }] };
    const report = importConfig(target, JSON.stringify(file));
    expect(report.ok).toBe(true);
    expect(report.profiles).toBe(1);
    expect(report.skipped).toHaveLength(4);
    expect(listProfiles(target, ND30).map((p) => p.id)).toEqual(["ok1"]);
  });

  it("file có phiên bản 1 vẫn đọc được khi thiếu các phần tuỳ chọn", () => {
    const report = importConfig(memoryStore(), JSON.stringify({ format: CONFIG_FORMAT, version: 1 }));
    expect(report).toMatchObject({ ok: true, profiles: 0, presets: 0 });
  });

  it("mọi chuẩn đều có thể xuất", () => {
    expect(STANDARDS.length).toBeGreaterThan(1);
    expect(() => exportConfig(memoryStore())).not.toThrow();
  });
});
