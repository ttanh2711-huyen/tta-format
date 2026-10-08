import { bodyPlan, pageNumberPlan, pagePlan, stylePlan, tablePlan } from "./planText";

const noSkips = { empty: 0, table: 0, standardStyle: 0, list: 0, heading: 0 };

describe("pagePlan", () => {
  it("đã đúng thì bỏ qua", () => {
    expect(pagePlan({ sections: 1, mismatches: [], marginsText: "20/20/30/20 mm" })).toEqual({
      willRun: false,
      lines: [],
      skipReason: "trang đã đúng thiết lập (A4, lề 20/20/30/20 mm).",
    });
  });

  it("lệch thì nêu số phần và các điểm lệch, rút gọn khi quá dài", () => {
    const plan = pagePlan({ sections: 2, mismatches: ["a", "b", "c", "d", "e", "f"], marginsText: "20/20/30/15 mm" });
    expect(plan.willRun).toBe(true);
    expect(plan.lines[0]).toBe("Đặt khổ A4 dọc, lề 20/20/30/15 mm cho 2 phần.");
    expect(plan.lines.slice(1)).toEqual(["Đang lệch: a; b; c; d; e", "… và 1 điểm lệch khác."]);
  });
});

describe("stylePlan", () => {
  it("không tạo, không cập nhật thì bỏ qua", () => {
    expect(stylePlan({ toCreate: 0, toUpdate: 0 }, "ND30_")).toMatchObject({ willRun: false, skipReason: "bộ style ND30_* đã đúng." });
  });

  it("nêu số style tạo và cập nhật", () => {
    expect(stylePlan({ toCreate: 3, toUpdate: 2 }, "HD05_")).toEqual({
      willRun: true,
      lines: ["Cài bộ style HD05_*: tạo 3 style, cập nhật 2 style."],
    });
  });
});

describe("bodyPlan", () => {
  it("không có đoạn nào cần định dạng thì bỏ qua và nêu lý do", () => {
    const plan = bodyPlan({ targets: 0, skipped: { ...noSkips, standardStyle: 4, empty: 2 }, styleMissing: false }, undefined, "ND30_NoiDung");
    expect(plan.willRun).toBe(false);
    expect(plan.skipReason).toBe("không có đoạn nào cần định dạng. Bỏ qua 6 đoạn: 2 đoạn trống, 4 đoạn đã dùng style của bộ chuẩn.");
  });

  it("nêu số đoạn, style áp và phần bỏ qua", () => {
    const plan = bodyPlan({ targets: 10, skipped: { ...noSkips, table: 3 }, styleMissing: false }, undefined, "ND30_NoiDung");
    expect(plan.willRun).toBe(true);
    expect(plan.lines).toEqual(["Áp style ND30_NoiDung cho 10 đoạn.", "Bỏ qua 3 đoạn: 3 đoạn trong bảng."]);
  });

  it("chuẩn Exactly: nói rõ cách đặt giãn dòng và rủi ro mất bình luận", () => {
    const plan = bodyPlan({ targets: 5, skipped: noSkips, styleMissing: false }, 20, "HD05_NoiDung");
    expect(plan.lines).toContain("Giãn dòng Exactly 20 pt đặt trực tiếp trên từng đoạn bằng cách ghi lại đoạn; bình luận hoặc đánh dấu gắn trong đoạn có thể không còn.");
  });
});

describe("tablePlan", () => {
  it("không có bảng dữ liệu thì bỏ qua và nói về bảng bố cục được giữ nguyên", () => {
    const plan = tablePlan({ data: 0, layout: 2, nested: 0, paragraphsToStyle: 0, styleMissing: false }, "ND30_BangDuLieu");
    expect(plan).toEqual({ willRun: false, lines: [], skipReason: "không có bảng dữ liệu cần định dạng. Giữ nguyên 2 bảng không viền (đầu văn bản, nơi nhận, chữ ký)." });
  });

  it("bảng dữ liệu đã dùng style rồi thì bỏ qua", () => {
    expect(tablePlan({ data: 1, layout: 0, nested: 0, paragraphsToStyle: 0, styleMissing: false }, "ND30_BangDuLieu").willRun).toBe(false);
  });

  it("nêu số bảng, số đoạn, và bảng giữ nguyên", () => {
    const plan = tablePlan({ data: 2, layout: 1, nested: 1, paragraphsToStyle: 14, styleMissing: false }, "ND30_BangDuLieu");
    expect(plan.willRun).toBe(true);
    expect(plan.lines).toEqual(["Áp style ND30_BangDuLieu cho 14 đoạn trong 2 bảng có viền.", "Giữ nguyên 1 bảng không viền và 1 bảng lồng nhau."]);
  });
});

describe("pageNumberPlan", () => {
  const none = { setDifferentFirst: undefined, writePrimary: false, clearFirst: false };

  it("đã đúng thì bỏ qua", () => {
    expect(pageNumberPlan([none], { sizePt: 14, hideFirst: true })).toEqual({ willRun: false, lines: [], skipReason: "số trang đã đúng (giữa trang, 14 pt, trang 1 không hiện số)." });
    expect(pageNumberPlan([none], { sizePt: 14, hideFirst: false }).skipReason).toBe("số trang đã đúng (giữa trang, 14 pt).");
  });

  it("nêu từng việc cần làm", () => {
    const plan = pageNumberPlan([{ setDifferentFirst: true, writePrimary: true, clearFirst: true }, { setDifferentFirst: undefined, writePrimary: true, clearFirst: false }], { sizePt: 14, hideFirst: true });
    expect(plan).toEqual({
      willRun: true,
      lines: ['Ghi số trang (căn giữa, 14 pt) vào header của 2 phần.', 'Đặt lại "trang đầu khác biệt" cho 1 phần.', "Xoá số trang đang nằm ở header trang đầu (trang 1 không hiện số)."],
    });
  });
});
