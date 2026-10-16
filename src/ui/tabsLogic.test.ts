import { memoryStore } from "../settings/storage";
import { STANDARDS } from "../standards";
import { loadActiveSection, loadActiveTab, nextTabIndex, saveActiveSection, saveActiveTab } from "./tabsLogic";

describe("phím di chuyển giữa các tab", () => {
  it.each([
    [0, "ArrowRight", 1],
    [3, "ArrowRight", 0],
    [0, "ArrowLeft", 3],
    [2, "ArrowLeft", 1],
    [2, "Home", 0],
    [1, "End", 3],
    [1, "ArrowDown", 2],
    [0, "ArrowUp", 3],
  ])("tab %i, phím %s đến tab %i (4 tab)", (current, key, expected) => {
    expect(nextTabIndex(current, 4, key)).toBe(expected);
  });

  it("phím khác hoặc không có tab thì không đổi", () => {
    expect(nextTabIndex(1, 4, "Enter")).toBeNull();
    expect(nextTabIndex(1, 4, "a")).toBeNull();
    expect(nextTabIndex(0, 0, "ArrowRight")).toBeNull();
  });
});

describe("mục con đang mở của từng tab", () => {
  it("mỗi tab nhớ riêng; mục không còn thì về mục đầu", () => {
    const store = memoryStore();
    expect(loadActiveSection(store, "check", ["a", "b"])).toBe("a");
    saveActiveSection(store, "check", "b");
    saveActiveSection(store, "agency", "x");
    expect(loadActiveSection(store, "check", ["a", "b"])).toBe("b");
    expect(loadActiveSection(store, "agency", ["y", "z"])).toBe("y");
  });
});

describe("tab đang chọn", () => {
  const ids = ["x", "y", "z"];

  it("lấy tab đã lưu nếu còn, không thì tab đầu", () => {
    const store = memoryStore();
    expect(loadActiveTab(store, ids)).toBe("x");
    saveActiveTab(store, "z");
    expect(loadActiveTab(store, ids)).toBe("z");
    saveActiveTab(store, "khong-con");
    expect(loadActiveTab(store, ids)).toBe("x");
  });
});

describe("nhãn tab của từng chuẩn", () => {
  it("mỗi chuẩn có nhãn tab riêng, không trùng", () => {
    const labels = STANDARDS.map((standard) => standard.tabLabel);
    expect(labels.every((label) => label.trim() !== "")).toBe(true);
    expect(new Set(labels).size).toBe(labels.length);
    expect(labels).toEqual(["Nghị định 30", "Hướng dẫn 05"]);
  });
});
