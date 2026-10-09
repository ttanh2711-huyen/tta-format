import { summarizeIssues } from "./engine";
import type { Issue } from "./types";

export interface ReportHeader {
  standardName: string;
  /** Thời điểm kiểm tra, để in vào báo cáo. */
  checkedAt: Date;
}

function where(issue: Issue): string {
  return issue.location.kind === "page" ? `Trang, phần ${issue.location.section}` : `Đoạn "${issue.location.excerpt}"`;
}

/** Báo cáo lỗi dạng chữ thuần (UTF-8), mỗi lỗi một mục: mức độ, mô tả, vị trí, điều khoản và cách sửa nếu có. */
export function formatReport(issues: readonly Issue[], header: ReportHeader): string {
  const counts = summarizeIssues(issues);
  const fixable = issues.filter((issue) => issue.fix).length;
  const when = header.checkedAt.toLocaleString("vi-VN");
  const lines = [`BÁO CÁO KIỂM TRA VĂN BẢN THEO CHUẨN ${header.standardName.toUpperCase()}`, `Kiểm tra lúc: ${when}`, `Tổng: ${counts.errors} lỗi, ${counts.warnings} cảnh báo; ${fixable} mục sửa được tự động.`, ""];
  if (issues.length === 0) {
    lines.push("Không thấy lỗi.");
    return lines.join("\n");
  }
  issues.forEach((issue, i) => {
    lines.push(`${i + 1}. [${issue.severity === "error" ? "Lỗi" : "Cảnh báo"}] ${issue.message}`);
    lines.push(`   Vị trí: ${where(issue)}`);
    lines.push(`   Căn cứ: ${issue.reference}`);
    if (issue.actual !== undefined || issue.expected !== undefined) lines.push(`   Hiện có: ${issue.actual ?? "-"} | Cần: ${issue.expected ?? "-"}`);
    lines.push(`   Cách sửa: ${issue.fix ? issue.fix.label : "cần sửa thủ công"}`);
  });
  return lines.join("\n");
}
