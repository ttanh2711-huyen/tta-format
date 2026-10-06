import { loadDefaultStandard, saveDefaultStandard } from "./defaultStandard";
import { memoryStore } from "./storage";

describe("chuẩn mặc định cho tài liệu mới", () => {
  it("chưa lưu thì dùng chuẩn đầu tiên trong danh sách (Chính quyền)", () => {
    expect(loadDefaultStandard(memoryStore()).id).toBe("nd30");
  });

  it("lưu rồi đọc lại đúng chuẩn", () => {
    const store = memoryStore();
    saveDefaultStandard(store, "hd05");
    expect(loadDefaultStandard(store).id).toBe("hd05");
  });

  it("giá trị lưu không còn hợp lệ thì dùng chuẩn đầu tiên", () => {
    const store = memoryStore();
    store.set("nd30.defaultStandard", "khong-co");
    expect(loadDefaultStandard(store).id).toBe("nd30");
  });
});
