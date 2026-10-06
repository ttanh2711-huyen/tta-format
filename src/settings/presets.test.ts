import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { buildRegistry, initialState, setOption, type FormatState } from "./formatOptions";
import {
  applyPreset,
  clearDefaults,
  deletePreset,
  duplicatePreset,
  listPresets,
  loadDefaults,
  renamePreset,
  saveDefaults,
  savePreset,
} from "./presets";
import { memoryStore } from "./storage";

const nd = buildRegistry(ND30);
const hd = buildRegistry(HD05);

function withFont(size: number, confirmed = false): FormatState {
  const result = setOption(nd, initialState(nd), "fontSize", size, { confirmed });
  if (result.status !== "applied") throw new Error("kỳ vọng applied");
  return result.state;
}

describe("preset có sẵn", () => {
  it("mỗi chuẩn có preset riêng theo tên đã khai báo, giá trị là mặc định của chuẩn", () => {
    const ndList = listPresets(memoryStore(), ND30);
    expect(ndList).toHaveLength(1);
    expect(ndList[0]).toMatchObject({ name: "Theo mẫu công văn của tôi", standardId: "nd30", builtIn: true });
    expect(ndList[0].values).toEqual(initialState(nd).values);
    expect(listPresets(memoryStore(), HD05)[0]).toMatchObject({ name: "Theo HD 05 (mặc định đề xuất)", standardId: "hd05" });
  });

  it("không xoá, không đổi tên preset có sẵn", () => {
    const store = memoryStore();
    const builtIn = listPresets(store, ND30)[0];
    expect(deletePreset(store, builtIn.id)).toEqual({ ok: false, reason: expect.stringContaining("có sẵn") });
    expect(renamePreset(store, builtIn.id, "Tên khác")).toEqual({ ok: false, reason: expect.stringContaining("có sẵn") });
    expect(listPresets(store, ND30)).toHaveLength(1);
  });
});

describe("lưu, đổi tên, nhân bản, xoá preset", () => {
  it("lưu preset của người dùng, còn nguyên khi đọc lại từ cùng kho lưu", () => {
    const store = memoryStore();
    const saved = savePreset(store, ND30, "Báo cáo dài", withFont(13));
    expect(saved.ok).toBe(true);
    const list = listPresets(store, ND30);
    expect(list.map((p) => p.name)).toEqual(["Theo mẫu công văn của tôi", "Báo cáo dài"]);
    expect(list[1].values.fontSize).toBe(13);
    expect(list[1].builtIn).toBeUndefined();
  });

  it("mỗi preset thuộc một chuẩn: chuẩn kia không thấy", () => {
    const store = memoryStore();
    savePreset(store, ND30, "Báo cáo dài", withFont(13));
    expect(listPresets(store, HD05).map((p) => p.name)).toEqual(["Theo HD 05 (mặc định đề xuất)"]);
  });

  it("tên rỗng hoặc trùng tên trong cùng chuẩn bị từ chối", () => {
    const store = memoryStore();
    expect(savePreset(store, ND30, "   ", withFont(13))).toEqual({ ok: false, reason: expect.stringContaining("tên") });
    savePreset(store, ND30, "Công văn thường", withFont(13));
    expect(savePreset(store, ND30, "công văn thường", withFont(14))).toEqual({ ok: false, reason: expect.stringContaining("trùng") });
    expect(savePreset(store, ND30, "Theo mẫu công văn của tôi", withFont(14)).ok).toBe(false);
    expect(savePreset(store, HD05, "Công văn thường", initialState(hd)).ok).toBe(true);
  });

  it("đổi tên và xoá preset của người dùng", () => {
    const store = memoryStore();
    const saved = savePreset(store, ND30, "A", withFont(13));
    if (!saved.ok) throw new Error("kỳ vọng ok");
    expect(renamePreset(store, saved.preset.id, "B").ok).toBe(true);
    expect(listPresets(store, ND30).map((p) => p.name)).toContain("B");
    expect(deletePreset(store, saved.preset.id).ok).toBe(true);
    expect(listPresets(store, ND30)).toHaveLength(1);
  });

  it("nhân bản preset, kể cả preset có sẵn, thêm hậu tố bản sao", () => {
    const store = memoryStore();
    const builtIn = listPresets(store, ND30)[0];
    const copy = duplicatePreset(store, ND30, builtIn.id);
    if (!copy.ok) throw new Error("kỳ vọng ok");
    expect(copy.preset.name).toBe("Theo mẫu công văn của tôi (bản sao)");
    expect(copy.preset.builtIn).toBeUndefined();
    const second = duplicatePreset(store, ND30, builtIn.id);
    if (!second.ok) throw new Error("kỳ vọng ok");
    expect(second.preset.name).toBe("Theo mẫu công văn của tôi (bản sao 2)");
  });

  it("mã preset không trùng nhau kể cả sau khi xoá", () => {
    const store = memoryStore();
    const a = savePreset(store, ND30, "A", withFont(13));
    const b = savePreset(store, ND30, "B", withFont(13));
    if (!a.ok || !b.ok) throw new Error("kỳ vọng ok");
    deletePreset(store, a.preset.id);
    const c = savePreset(store, ND30, "C", withFont(13));
    if (!c.ok) throw new Error("kỳ vọng ok");
    expect(new Set([b.preset.id, c.preset.id]).size).toBe(2);
  });

  it("kho lưu hỏng thì coi như rỗng, không ném lỗi", () => {
    const store = memoryStore();
    store.set("nd30.presets", "{không phải json");
    expect(listPresets(store, ND30)).toHaveLength(1);
    expect(savePreset(store, ND30, "A", withFont(13)).ok).toBe(true);
  });
});

describe("áp lại preset", () => {
  it("trả về trạng thái đúng giá trị đã lưu và ngoại lệ chủ ý", () => {
    const store = memoryStore();
    savePreset(store, ND30, "Chữ lớn", withFont(16, true));
    const preset = listPresets(store, ND30)[1];
    const { state, ignored } = applyPreset(nd, preset);
    expect(state.values.fontSize).toBe(16);
    expect(state.exceptions).toEqual(["fontSize"]);
    expect(ignored).toEqual([]);
  });

  it("giá trị lỗi thời bị bỏ qua và có thông báo", () => {
    const store = memoryStore();
    store.set("nd30.presets", JSON.stringify([{ id: "u1", name: "Cũ", standardId: "nd30", values: { fontSize: 13, tuyChonCu: 7 }, exceptions: [] }]));
    const preset = listPresets(store, ND30)[1];
    const { state, ignored } = applyPreset(nd, preset);
    expect(state.values.fontSize).toBe(13);
    expect(ignored).toHaveLength(1);
    expect(ignored[0]).toContain("tuyChonCu");
  });
});

describe("lưu làm mặc định", () => {
  it("chưa lưu thì dùng mặc định của chuẩn", () => {
    const { state, ignored } = loadDefaults(memoryStore(), nd, "nd30");
    expect(state).toEqual(initialState(nd));
    expect(ignored).toEqual([]);
  });

  it("lưu rồi đọc lại đúng giá trị, mỗi chuẩn một bản", () => {
    const store = memoryStore();
    saveDefaults(store, "nd30", withFont(13));
    expect(loadDefaults(store, nd, "nd30").state.values.fontSize).toBe(13);
    expect(loadDefaults(store, hd, "hd05").state).toEqual(initialState(hd));
  });

  it("xoá mặc định đã lưu thì quay về mặc định của chuẩn", () => {
    const store = memoryStore();
    saveDefaults(store, "nd30", withFont(13));
    clearDefaults(store, "nd30");
    expect(loadDefaults(store, nd, "nd30").state.values.fontSize).toBe(14);
  });
});
