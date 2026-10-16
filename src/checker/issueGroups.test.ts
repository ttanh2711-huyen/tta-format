import { groupIssues, groupTitle } from "./issueGroups";
import type { Issue } from "./types";

const issue = (ruleId: string, index: number, over: Partial<Issue> = {}): Issue => ({
  ruleId,
  severity: "error",
  message: `${ruleId} ở đoạn ${index}`,
  reference: "ref",
  location: { kind: "paragraph", index, excerpt: `Đoạn ${index}` },
  ...over,
});

describe("gom lỗi kiểm tra", () => {
  it("lỗi cùng quy tắc vào một nhóm, đếm đủ số chỗ và gom cách sửa", () => {
    const fix = { kind: "font" as const, index: 1, sizePt: 14, label: "Cỡ 14" };
    const [category] = groupIssues([issue("font.size", 1, { fix }), issue("font.size", 2), issue("font.size", 3, { fix: { ...fix, index: 3 } })]);
    expect(category.label).toBe("Phông và cỡ chữ");
    expect(category.count).toBe(3);
    expect(category.groups).toHaveLength(1);
    expect(category.groups[0].title).toBe("Cỡ chữ");
    expect(category.groups[0].issues.map((i) => i.location)).toHaveLength(3);
    expect(category.groups[0].fixes).toHaveLength(2);
  });

  it("khu vực xếp theo thứ tự trong văn bản, không theo thứ tự lỗi xuất hiện", () => {
    const categories = groupIssues([issue("cap.sentenceStart", 5), issue("footer.hoTen.bold", 9), issue("page.margin.left", 0), issue("header.date.italic", 1), issue("font.name", 2)]);
    expect(categories.map((c) => c.id)).toEqual(["page", "font", "header", "footer", "cap"]);
  });

  it("trong một khu vực, nhóm có lỗi đứng trước nhóm chỉ có cảnh báo", () => {
    const [category] = groupIssues([issue("paragraph.spacing", 1, { severity: "warning" }), issue("paragraph.firstLineIndent", 2)]);
    expect(category.groups.map((g) => g.ruleId)).toEqual(["paragraph.firstLineIndent", "paragraph.spacing"]);
  });

  it("nhóm lấy mức nặng nhất của các chỗ trong nhóm", () => {
    const [category] = groupIssues([issue("font.size", 1, { severity: "warning" }), issue("font.size", 2)]);
    expect(category.groups[0].severity).toBe("error");
  });

  it("quy tắc lạ vào khu vực Khác", () => {
    expect(groupIssues([issue("zzz.unknown", 0)]).map((c) => c.label)).toEqual(["Khác"]);
  });

  it("tên nhóm: bảng có sẵn, ghép thành phần với thuộc tính, không thì lấy câu báo lỗi", () => {
    expect(groupTitle("page.margin.left", [])).toBe("Lề trái");
    expect(groupTitle("footer.chucVu.size", [])).toBe("Chức vụ người ký: cỡ chữ");
    expect(groupTitle("urgency.bold", [])).toBe("Dấu mức độ khẩn: in đậm");
    expect(groupTitle("structure.phanChuong.upper", [issue("structure.phanChuong.upper", 0, { message: "Phần, chương phải in hoa." })])).toBe("Phần, chương phải in hoa.");
  });

  it("không có lỗi thì không có khu vực nào", () => {
    expect(groupIssues([])).toEqual([]);
  });
});
