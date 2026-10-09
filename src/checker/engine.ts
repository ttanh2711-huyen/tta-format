import type { StandardProfile } from "../standards/types";
import { checkPage } from "./pageRules";
import { checkParagraphs } from "./paragraphRules";
import type { CheckDocument, Issue } from "./types";

/**
 * Engine kiểm tra: đối chiếu tài liệu với bảng quy tắc của một chuẩn và trả về danh sách lỗi, lỗi về trang trước
 * rồi lỗi về đoạn theo thứ tự trong tài liệu. Thuần tuý, không gọi Word; các bộ quy tắc khác (khối đầu, khối cuối)
 * đăng ký thêm qua `extraRules`.
 */
export type Rule = (doc: CheckDocument, profile: StandardProfile) => Issue[];

/** Lỗi trang đứng trước lỗi đoạn; lỗi đoạn theo chỉ số. */
function order(issue: Issue): number {
  return issue.location.kind === "page" ? issue.location.section - 1_000_000 : issue.location.index;
}

export function runChecks(doc: CheckDocument, profile: StandardProfile, extraRules: readonly Rule[] = []): Issue[] {
  const pageIssues = doc.pages.flatMap((page, index) => checkPage(page, index + 1, profile));
  const rest = [checkParagraphs, ...extraRules].flatMap((rule) => rule(doc, profile));
  return [...pageIssues, ...rest].sort((a, b) => order(a) - order(b));
}

export interface IssueSummary {
  errors: number;
  warnings: number;
}

export function summarizeIssues(issues: readonly Issue[]): IssueSummary {
  return { errors: issues.filter((i) => i.severity === "error").length, warnings: issues.filter((i) => i.severity === "warning").length };
}
