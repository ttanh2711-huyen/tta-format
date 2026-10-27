import type { Fix, Issue } from "./types";

/**
 * Gom lỗi kiểm tra để hiển thị gọn: lỗi cùng quy tắc (cùng `ruleId`) thành một nhóm, nhóm xếp theo khu vực của văn bản
 * (trang, phông, đoạn văn, khối đầu, căn cứ, khối cuối, viết hoa, thành phần khác). Mỗi nhóm có tên ngắn, số chỗ, mức
 * nặng nhất và các cách sửa để sửa cả nhóm một lần.
 */

export interface IssueGroup {
  ruleId: string;
  title: string;
  severity: Issue["severity"];
  issues: Issue[];
  fixes: Fix[];
}

export interface IssueCategory {
  id: string;
  label: string;
  groups: IssueGroup[];
  /** Tổng số chỗ của mọi nhóm trong khu vực. */
  count: number;
}

const CATEGORIES: ReadonlyArray<{ id: string; label: string; prefixes: readonly string[] }> = [
  { id: "page", label: "Trang", prefixes: ["page"] },
  { id: "font", label: "Phông và cỡ chữ", prefixes: ["font"] },
  { id: "paragraph", label: "Đoạn văn và bố cục", prefixes: ["paragraph", "list", "structure"] },
  { id: "header", label: "Khối đầu văn bản", prefixes: ["header", "title", "urgency"] },
  { id: "basis", label: "Căn cứ", prefixes: ["basis"] },
  { id: "footer", label: "Khối cuối văn bản", prefixes: ["footer"] },
  { id: "cap", label: "Viết hoa", prefixes: ["cap"] },
  { id: "extras", label: "Thành phần khác", prefixes: ["extras"] },
];
const OTHER = { id: "other", label: "Khác" };

const TITLES: Readonly<Record<string, string>> = {
  "page.paper": "Khổ giấy",
  "page.orientation": "Hướng giấy",
  "page.margin.top": "Lề trên",
  "page.margin.bottom": "Lề dưới",
  "page.margin.left": "Lề trái",
  "page.margin.right": "Lề phải",
  "page.headerDistance": "Khoảng cách số trang tới mép trên",
  "font.name": "Phông chữ",
  "font.size": "Cỡ chữ",
  "paragraph.align": "Căn đều hai bên",
  "paragraph.firstLineIndent": "Thụt dòng đầu",
  "paragraph.spacing": "Giãn đoạn",
  "paragraph.lineSpacing": "Giãn dòng",
  "paragraph.endsWithPeriod": "Dấu chấm cuối nội dung",
  "list.diem.order": "Thứ tự điểm a), b), c)",
  "structure.dieu.bold": "Điều in đậm",
  "structure.dieu.sequence": "Số thứ tự Điều",
  "cap.sentenceStart": "Viết hoa đầu câu",
  "cap.afterPeriod": "Viết hoa sau dấu chấm",
  "cap.properName": "Viết hoa tên riêng",
  "cap.documentName": "Viết hoa tên văn bản khi viện dẫn",
  "cap.reference": "Viết hoa Phần, Chương, Mục, Điều khi viện dẫn",
  "cap.referenceLower": "Viết thường khoản, điểm khi viện dẫn",
  "basis.dash": "Gạch đầu dòng của căn cứ",
  "basis.upright": "Căn cứ: chữ đứng",
  "basis.punctuation": "Dấu cuối mỗi căn cứ",
  "basis.verdict.bold": "Dòng chốt QUYẾT ĐỊNH, QUYẾT NGHỊ in đậm",
  "header.mixed": "Lẫn khối đầu của hai chuẩn",
  "header.otherStandard": "Khối đầu theo chuẩn khác",
  "header.tieuNguChinhQuyen": "Tiêu ngữ không thuộc chuẩn",
  "header.coQuanBanHanh.missing": "Thiếu tên cơ quan ban hành",
  "header.star": "Dấu sao dưới tên cơ quan",
  "header.so.format": "Số và ký hiệu",
  "header.date.format": "Cách ghi địa danh, ngày tháng",
  "header.date.invalid": "Ngày tháng không tồn tại",
  "header.date.year": "Năm ghi đủ 4 chữ số",
  "title.dashes.missing": "Thiếu gạch nối dưới trích yếu",
  "title.dashes.count": "Số gạch nối dưới trích yếu",
  "footer.noiNhan.dash": "Gạch đầu dòng nơi nhận",
  "footer.noiNhan.punctuation": "Dấu câu cuối mục nơi nhận",
  "footer.noiNhan.format": "Chữ \"Nơi nhận\"",
  "footer.quyenHan.format": "Quyền hạn ký",
  "footer.hoTen.hocHam": "Học hàm, học vị trong họ tên",
  "footer.hoTen.case": "Họ tên in thường",
};

/** Tên thành phần (đoạn giữa của mã quy tắc) và thuộc tính (đoạn cuối) để dựng tên nhóm cho mã chưa có trong bảng. */
const SUBJECTS: Readonly<Record<string, string>> = {
  coQuanBanHanh: "Tên cơ quan ban hành",
  date: "Địa danh, ngày tháng",
  kinhGui: "Kính gửi",
  theLoai: "Thể loại",
  trichYeu: "Trích yếu",
  chucVu: "Chức vụ người ký",
  hoTen: "Họ tên người ký",
  noiNhan: "Danh sách nơi nhận",
  draft: "Chữ \"Dự thảo\"",
  noPublic: "Phạm vi lưu hành",
  verdict: "Chữ quyết nghị",
  urgency: "Dấu mức độ khẩn",
  basis: "Căn cứ",
};
const ATTRIBUTES: Readonly<Record<string, string>> = {
  size: "cỡ chữ",
  bold: "in đậm",
  italic: "in nghiêng",
  upper: "in hoa",
  text: "viết đúng",
  punctuation: "dấu câu",
};

const SEVERITY_RANK: Record<Issue["severity"], number> = { error: 0, warning: 1 };

export function groupTitle(ruleId: string, issues: readonly Issue[]): string {
  const known = TITLES[ruleId];
  if (known) return known;
  const parts = ruleId.split(".");
  const attribute = ATTRIBUTES[parts[parts.length - 1]];
  const subject = SUBJECTS[parts.length > 2 ? parts[parts.length - 2] : parts[0]];
  if (subject && attribute) return `${subject}: ${attribute}`;
  return issues[0]?.message ?? ruleId;
}

function categoryOf(ruleId: string): { id: string; label: string } {
  const head = ruleId.split(".")[0];
  return CATEGORIES.find((category) => category.prefixes.includes(head)) ?? OTHER;
}

/** Gom và xếp lỗi theo khu vực, rồi theo mức nặng (lỗi trước cảnh báo), rồi theo thứ tự xuất hiện. */
export function groupIssues(issues: readonly Issue[]): IssueCategory[] {
  const groups = new Map<string, IssueGroup>();
  for (const issue of issues) {
    let group = groups.get(issue.ruleId);
    if (!group) {
      group = { ruleId: issue.ruleId, title: "", severity: issue.severity, issues: [], fixes: [] };
      groups.set(issue.ruleId, group);
    }
    group.issues.push(issue);
    if (issue.fix) group.fixes.push(issue.fix);
    if (SEVERITY_RANK[issue.severity] < SEVERITY_RANK[group.severity]) group.severity = issue.severity;
  }
  const order = [...CATEGORIES.map((category) => category.id), OTHER.id];
  const categories = new Map<string, IssueCategory>();
  for (const group of groups.values()) {
    group.title = groupTitle(group.ruleId, group.issues);
    const { id, label } = categoryOf(group.ruleId);
    const category = categories.get(id) ?? { id, label, groups: [], count: 0 };
    category.groups.push(group);
    category.count += group.issues.length;
    categories.set(id, category);
  }
  for (const category of categories.values()) category.groups.sort((a, b) => SEVERITY_RANK[a.severity] - SEVERITY_RANK[b.severity]);
  return [...categories.values()].sort((a, b) => order.indexOf(a.id) - order.indexOf(b.id));
}
