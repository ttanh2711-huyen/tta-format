import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { cleanRecipientLine, listAgencies, removeAgency, saveAgency, saveRecipients } from "./agencies";
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

describe("danh bạ nơi nhận", () => {
  it("làm sạch dòng nơi nhận: bỏ gạch đầu dòng, dấu cuối dòng, khoảng trắng thừa", () => {
    expect(cleanRecipientLine("  - Sở  Nội vụ;")).toBe("Sở Nội vụ");
    expect(cleanRecipientLine("– Văn phòng Chính phủ.")).toBe("Văn phòng Chính phủ");
    expect(cleanRecipientLine("   ")).toBe("");
  });

  it("lưu nhiều dòng một lần, bỏ dòng trống, 'Lưu' và 'Như trên', không trùng", () => {
    const store = memoryStore();
    const result = saveRecipients(store, ND30, ["Sở Nội vụ", "- Sở Tài chính;", "", "Lưu: VT", "Như trên", "sở nội vụ"]);
    expect(result.added).toBe(2);
    expect(listAgencies(store, ND30, "recipients")).toEqual(["Sở Nội vụ", "Sở Tài chính"]);
    expect(saveRecipients(store, ND30, ["Sở Nội vụ", "Sở Y tế"]).added).toBe(1);
    expect(listAgencies(store, ND30, "recipients")).toEqual(["Sở Nội vụ", "Sở Tài chính", "Sở Y tế"]);
  });

  it("mỗi chuẩn một danh bạ riêng; xoá được từng mục", () => {
    const store = memoryStore();
    saveRecipients(store, ND30, ["A", "B"]);
    expect(listAgencies(store, HD05, "recipients")).toEqual([]);
    removeAgency(store, ND30, "A", "recipients");
    expect(listAgencies(store, ND30, "recipients")).toEqual(["B"]);
  });

  it("danh bạ độc lập với danh sách tên cơ quan ban hành", () => {
    const store = memoryStore();
    saveAgency(store, ND30, "Chi cục Thủy lợi");
    saveRecipients(store, ND30, ["Sở Nội vụ"]);
    expect(listAgencies(store, ND30)).toEqual(["Chi cục Thủy lợi"]);
  });
});
