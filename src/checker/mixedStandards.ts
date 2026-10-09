import { STANDARDS } from "../standards";
import type { StandardProfile } from "../standards/types";
import { matchTitle, paragraphIssue } from "./common";
import type { CheckDocument, Issue } from "./types";

/**
 * Phát hiện tài liệu mang tiêu đề của chuẩn khác (quốc hiệu của Chính quyền trong văn bản Đảng và ngược lại).
 * Có cả hai tiêu đề thì cảnh báo trộn hai chuẩn; chỉ có tiêu đề của chuẩn kia thì báo lỗi. Không có cách sửa
 * tự động vì đổi chuẩn là thao tác riêng có xác nhận.
 */
export function checkMixedStandards(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const own = doc.paragraphs.some((p) => p.inTable && matchTitle(p.resolved.text, profile.identifyingTitle) !== "no");
  const issues: Issue[] = [];
  for (const other of STANDARDS) {
    if (other.id === profile.id) continue;
    for (const p of doc.paragraphs) {
      if (!p.inTable || matchTitle(p.resolved.text, other.identifyingTitle) === "no") continue;
      issues.push(
        paragraphIssue(p, {
          ruleId: own ? "header.mixed" : "header.otherStandard",
          severity: own ? "warning" : "error",
          message: own
            ? `Tài liệu có cả tiêu đề của chuẩn ${profile.name} và của chuẩn ${other.name}; hãy giữ một chuẩn.`
            : `Đây là tiêu đề của chuẩn ${other.name}, không phải chuẩn ${profile.name} đang chọn; chọn đúng chuẩn hoặc đổi tiêu đề.`,
          reference: other.identifyingTitle,
          actual: p.resolved.text.trim(),
          expected: profile.identifyingTitle,
        }),
      );
    }
  }
  return issues;
}
