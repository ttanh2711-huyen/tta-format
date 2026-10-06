import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import {
  classifyParagraph,
  planFormatBody,
  summarizeFormat,
  type ParagraphInfo,
} from "./paragraphClassifier";

function para(partial: Partial<ParagraphInfo> = {}): ParagraphInfo {
  return {
    text: "Thực hiện Thông báo số 458/TB-VPUB ngày 02/10/2026 của Văn phòng Ủy ban nhân dân thành phố.",
    styleName: "Normal",
    styleBuiltIn: "Normal",
    tableNestingLevel: 0,
    isListItem: false,
    ...partial,
  };
}

describe("classifyParagraph", () => {
  it("đoạn văn bản thường dùng style Normal được áp dụng", () => {
    expect(classifyParagraph(para())).toEqual({ apply: true });
  });

  it("đoạn dùng style tự đặt của người dùng vẫn được áp dụng", () => {
    expect(classifyParagraph(para({ styleName: "Văn bản của tôi", styleBuiltIn: "Other" }))).toEqual({ apply: true });
  });

  it("bỏ qua đoạn trống hoặc chỉ có khoảng trắng", () => {
    expect(classifyParagraph(para({ text: "" }))).toEqual({ apply: false, reason: "empty" });
    expect(classifyParagraph(para({ text: "  \t " }))).toEqual({ apply: false, reason: "empty" });
  });

  it("bỏ qua đoạn nằm trong bảng, kể cả bảng lồng nhau", () => {
    expect(classifyParagraph(para({ tableNestingLevel: 1 }))).toEqual({ apply: false, reason: "table" });
    expect(classifyParagraph(para({ tableNestingLevel: 2 }))).toEqual({ apply: false, reason: "table" });
  });

  it("bỏ qua đoạn đã dùng style của bất kỳ chuẩn nào (ND30_ hoặc HD05_)", () => {
    expect(classifyParagraph(para({ styleName: "ND30_Dieu" }))).toEqual({ apply: false, reason: "standardStyle" });
    expect(classifyParagraph(para({ styleName: "HD05_NoiDung" }))).toEqual({ apply: false, reason: "standardStyle" });
    expect(classifyParagraph(para({ styleName: "ND30_NoiDung" }))).toEqual({ apply: false, reason: "standardStyle" });
  });

  it("bỏ qua tiêu đề có sẵn của Word, không phụ thuộc ngôn ngữ giao diện", () => {
    expect(classifyParagraph(para({ styleName: "Tiêu đề 1", styleBuiltIn: "Heading1" }))).toEqual({
      apply: false,
      reason: "heading",
    });
    expect(classifyParagraph(para({ styleName: "Title", styleBuiltIn: "Title" }))).toEqual({ apply: false, reason: "heading" });
  });

  it("bỏ qua mục của danh sách tự động để không làm hỏng đánh số", () => {
    expect(classifyParagraph(para({ isListItem: true }))).toEqual({ apply: false, reason: "list" });
  });

  it("đoạn trống nằm trong bảng ghi lý do trống trước (không quan trọng với kết quả áp dụng)", () => {
    expect(classifyParagraph(para({ text: "", tableNestingLevel: 1 })).apply).toBe(false);
  });
});

describe("planFormatBody", () => {
  const paragraphs: ParagraphInfo[] = [
    para({ text: "Kính gửi: Sở Nông nghiệp", styleName: "Normal" }),
    para({ text: "" }),
    para({ text: "Ô trong bảng", tableNestingLevel: 1 }),
    para({ text: "Điều 1. Phạm vi", styleName: "ND30_Dieu" }),
    para({ text: "Nội dung thứ hai" }),
    para({ text: "Mục tự động", isListItem: true }),
    para({ text: "Tiêu đề", styleName: "Heading 1", styleBuiltIn: "Heading1" }),
  ];

  it("trả về chỉ số các đoạn sẽ áp dụng và đếm từng lý do bỏ qua", () => {
    const plan = planFormatBody(paragraphs, ND30);
    expect(plan.targets).toEqual([0, 4]);
    expect(plan.skipped).toEqual({ empty: 1, table: 1, standardStyle: 1, list: 1, heading: 1 });
    expect(plan.styleName).toBe("ND30_NoiDung");
  });

  it("tên style đích theo chuẩn", () => {
    expect(planFormatBody([para()], HD05).styleName).toBe("HD05_NoiDung");
  });

  it("danh sách rỗng thì không có gì để làm", () => {
    expect(planFormatBody([], ND30)).toEqual({ targets: [], skipped: { empty: 0, table: 0, standardStyle: 0, list: 0, heading: 0 }, styleName: "ND30_NoiDung" });
  });
});

describe("summarizeFormat", () => {
  const skipped = { empty: 2, table: 3, standardStyle: 1, list: 0, heading: 0 };

  it("báo số đoạn đã định dạng và số đoạn bỏ qua kèm lý do", () => {
    const result = summarizeFormat({ applied: 10, skipped, notes: [] });
    expect(result.kind).toBe("ok");
    expect(result.text).toBe("Đã định dạng 10 đoạn. Bỏ qua 6 đoạn: 2 đoạn trống, 3 đoạn trong bảng, 1 đoạn đã dùng style của bộ chuẩn.");
  });

  it("không nhắc phần bỏ qua khi không bỏ qua đoạn nào", () => {
    const result = summarizeFormat({ applied: 4, skipped: { empty: 0, table: 0, standardStyle: 0, list: 0, heading: 0 }, notes: [] });
    expect(result.text).toBe("Đã định dạng 4 đoạn.");
  });

  it("không có đoạn nào để định dạng là cảnh báo", () => {
    const result = summarizeFormat({ applied: 0, skipped, notes: [] });
    expect(result.kind).toBe("warning");
    expect(result.text).toContain("Không có đoạn nào để định dạng");
  });

  it("ghi chú bổ sung được nối vào cuối và đổi mức thành cảnh báo khi có ghi chú cảnh báo", () => {
    const result = summarizeFormat({ applied: 3, skipped: { empty: 0, table: 0, standardStyle: 0, list: 0, heading: 0 }, notes: ["Giãn dòng Exactly chưa đặt được."], warn: true });
    expect(result.kind).toBe("warning");
    expect(result.text).toBe("Đã định dạng 3 đoạn. Giãn dòng Exactly chưa đặt được.");
  });
});
