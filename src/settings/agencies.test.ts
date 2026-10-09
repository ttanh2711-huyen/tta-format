import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { listAgencies, removeAgency, saveAgency } from "./agencies";
import { memoryStore } from "./storage";

describe("danh sách cơ quan đã lưu", () => {
  it("ban đầu rỗng", () => {
    expect(listAgencies(memoryStore(), ND30)).toEqual([]);
  });

  it("lưu, đọc lại, chuẩn hoá khoảng trắng", () => {
    const store = memoryStore();
    expect(saveAgency(store, ND30, "  Chi cục   Thủy lợi ")).toEqual({ ok: true });
    expect(listAgencies(store, ND30)).toEqual(["Chi cục Thủy lợi"]);
  });

  it("không lưu tên rỗng, không lưu trùng (không phân biệt hoa thường)", () => {
    const store = memoryStore();
    expect(saveAgency(store, ND30, "   ")).toEqual({ ok: false, reason: expect.stringContaining("tên") });
    saveAgency(store, ND30, "Chi cục Thủy lợi");
    expect(saveAgency(store, ND30, "CHI CỤC THỦY LỢI")).toEqual({ ok: false, reason: expect.stringContaining("đã có") });
    expect(listAgencies(store, ND30)).toHaveLength(1);
  });

  it("mỗi chuẩn một danh sách riêng", () => {
    const store = memoryStore();
    saveAgency(store, ND30, "Chi cục Thủy lợi");
    expect(listAgencies(store, HD05)).toEqual([]);
  });

  it("xoá một cơ quan", () => {
    const store = memoryStore();
    saveAgency(store, ND30, "A");
    saveAgency(store, ND30, "B");
    removeAgency(store, ND30, "A");
    expect(listAgencies(store, ND30)).toEqual(["B"]);
  });

  it("kho hỏng thì coi như rỗng", () => {
    const store = memoryStore();
    store.set("nd30.agencies.nd30", "{hỏng");
    expect(listAgencies(store, ND30)).toEqual([]);
    store.set("nd30.agencies.nd30", JSON.stringify([1, "A", null]));
    expect(listAgencies(store, ND30)).toEqual(["A"]);
  });

  it("danh sách chữ viết tắt cơ quan tách riêng với danh sách tên cơ quan", () => {
    const store = memoryStore();
    saveAgency(store, HD05, "Ban Tổ chức");
    saveAgency(store, HD05, "BTCTU", "abbrs");
    expect(listAgencies(store, HD05)).toEqual(["Ban Tổ chức"]);
    expect(listAgencies(store, HD05, "abbrs")).toEqual(["BTCTU"]);
    removeAgency(store, HD05, "BTCTU", "abbrs");
    expect(listAgencies(store, HD05, "abbrs")).toEqual([]);
    expect(listAgencies(store, HD05)).toEqual(["Ban Tổ chức"]);
  });
});
