import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { counterKey, lastNumber, listCounters, nextNumber, normalizeTypeKey, recordNumber, removeCounter, setLastNumber, yearOf } from "./soVanBan";
import { memoryStore } from "./storage";

describe("sổ số văn bản", () => {
  it("chưa có số nào thì số tiếp theo là 1", () => {
    const store = memoryStore();
    expect(lastNumber(store, ND30, 2026, "QĐ")).toBe(0);
    expect(nextNumber(store, ND30, 2026, "QĐ")).toBe(1);
  });

  it("ghi số đã dùng thì số tiếp theo tăng lên; ghi số nhỏ hơn không làm lùi", () => {
    const store = memoryStore();
    recordNumber(store, ND30, 2026, "QĐ", 7);
    expect(nextNumber(store, ND30, 2026, "QĐ")).toBe(8);
    recordNumber(store, ND30, 2026, "QĐ", 3);
    expect(lastNumber(store, ND30, 2026, "QĐ")).toBe(7);
    recordNumber(store, ND30, 2026, "QĐ", 12);
    expect(nextNumber(store, ND30, 2026, "QĐ")).toBe(13);
  });

  it("đếm riêng theo loại, theo năm và theo chuẩn", () => {
    const store = memoryStore();
    recordNumber(store, ND30, 2026, "QĐ", 5);
    expect(nextNumber(store, ND30, 2026, "TB")).toBe(1);
    expect(nextNumber(store, ND30, 2027, "QĐ")).toBe(1);
    expect(nextNumber(store, HD05, 2026, "QĐ")).toBe(1);
  });

  it("loại không có chữ viết tắt (công văn) dùng một bộ đếm chung; chữ viết tắt không phân biệt hoa thường và khoảng trắng", () => {
    const store = memoryStore();
    recordNumber(store, ND30, 2026, null, 4);
    expect(nextNumber(store, ND30, 2026, "")).toBe(5);
    expect(nextNumber(store, ND30, 2026, "  ")).toBe(5);
    recordNumber(store, ND30, 2026, "qđ", 2);
    expect(lastNumber(store, ND30, 2026, " QĐ ")).toBe(2);
    expect(normalizeTypeKey(" qyđ ")).toBe("QYĐ");
    expect(normalizeTypeKey(null)).toBe("-");
  });

  it("số không hợp lệ (0, âm, không nguyên) bị bỏ qua", () => {
    const store = memoryStore();
    for (const bad of [0, -3, 1.5, Number.NaN]) recordNumber(store, ND30, 2026, "QĐ", bad);
    expect(lastNumber(store, ND30, 2026, "QĐ")).toBe(0);
    expect(setLastNumber(store, ND30, 2026, "QĐ", -1)).toBe(false);
  });

  it("sửa tay số cuối cùng (kể cả giảm) và xoá bộ đếm", () => {
    const store = memoryStore();
    recordNumber(store, ND30, 2026, "QĐ", 20);
    expect(setLastNumber(store, ND30, 2026, "QĐ", 15)).toBe(true);
    expect(nextNumber(store, ND30, 2026, "QĐ")).toBe(16);
    removeCounter(store, ND30, 2026, "QĐ");
    expect(nextNumber(store, ND30, 2026, "QĐ")).toBe(1);
    expect(listCounters(store, ND30)).toEqual([]);
  });

  it("liệt kê các bộ đếm của chuẩn, năm mới nhất trước", () => {
    const store = memoryStore();
    recordNumber(store, ND30, 2025, "QĐ", 9);
    recordNumber(store, ND30, 2026, "TB", 4);
    recordNumber(store, ND30, 2026, "QĐ", 6);
    recordNumber(store, HD05, 2026, "NQ", 1);
    expect(listCounters(store, ND30)).toEqual([
      { year: 2026, typeKey: "QĐ", last: 6 },
      { year: 2026, typeKey: "TB", last: 4 },
      { year: 2025, typeKey: "QĐ", last: 9 },
    ]);
  });

  it("kho hỏng thì coi như chưa có số", () => {
    const store = memoryStore();
    store.set(counterKey(ND30, 2026, "QĐ"), "không phải số");
    expect(lastNumber(store, ND30, 2026, "QĐ")).toBe(0);
    store.set("nd30.so.index.nd30", "{hỏng");
    expect(listCounters(store, ND30)).toEqual([]);
  });

  it("lấy năm từ ô năm; không có thì dùng năm hiện tại", () => {
    expect(yearOf("2026", new Date(2030, 0, 1))).toBe(2026);
    expect(yearOf("", new Date(2030, 0, 1))).toBe(2030);
    expect(yearOf("26", new Date(2030, 0, 1))).toBe(2030);
    expect(yearOf("abcd", new Date(2030, 0, 1))).toBe(2030);
  });
});
