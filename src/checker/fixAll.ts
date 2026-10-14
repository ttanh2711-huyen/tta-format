import { applyFix } from "./fixes";
import type { CheckDocument, Fix, Issue } from "./types";

/**
 * "Sửa tất cả" (T15), phần thuần tuý: gom các cách sửa của danh sách lỗi thành một kế hoạch ngắn nhất (bỏ cách sửa
 * trùng, gộp các cách sửa cùng một đoạn hoặc cùng một phần), sắp xếp, mô tả để xem trước, và giả lập việc sửa
 * trên bản sao để biết còn lại lỗi nào. Không gọi Word.
 */

export interface FixPlanOptions {
  /** Có áp cách sửa thay chữ không; bỏ qua thì các lỗi đó vẫn nằm trong `manual`. Mặc định có. */
  includeText?: boolean;
  /** Bỏ qua lỗi của các quy tắc mà hàm này trả true (lỗi vẫn nằm trong `manual`). */
  skipRule?: (ruleId: string) => boolean;
}

export interface PlannedChange {
  label: string;
  /** Chỗ sửa, dạng chữ để hiện cho người dùng. */
  where: string;
  ruleIds: string[];
}

export interface FixPlan {
  /** Các cách sửa đã gộp, theo thứ tự áp dụng. */
  fixes: Fix[];
  /** Mô tả từng cách sửa, cùng thứ tự với `fixes`. */
  changes: PlannedChange[];
  /** Lỗi không được sửa trong lượt này: không có cách sửa, hoặc bị loại, hoặc hoãn sang lượt sau. */
  manual: Issue[];
}

const KIND_ORDER: Record<Fix["kind"], number> = { page: 0, font: 1, paragraph: 2, replace: 3, edits: 3, exact: 4, insertAfter: 5 };

function targetOf(fix: Fix): string {
  return fix.kind === "page" ? `page:${fix.section}` : `para:${fix.index}`;
}

/** Gộp hai cách sửa cùng loại, cùng chỗ; undefined nếu loại này không gộp được. */
function merge(a: Fix, b: Fix): Fix | undefined {
  if (a.kind !== b.kind) return undefined;
  const label = a.label === b.label ? a.label : `${a.label}; ${b.label}`;
  if (a.kind === "page" && b.kind === "page") return { ...a, ...b, marginsMm: { ...a.marginsMm, ...b.marginsMm }, label };
  if (a.kind === "font" && b.kind === "font") return { ...a, ...b, label };
  if (a.kind === "paragraph" && b.kind === "paragraph") return { ...a, ...b, label };
  if (a.kind === "exact" && b.kind === "exact") return { ...b, label };
  // Các chỗ thay chữ không chồng nhau trên một đoạn áp được nối tiếp nhau.
  if (a.kind === "edits" && b.kind === "edits") return { ...a, edits: [...a.edits, ...b.edits], label };
  return undefined;
}

function whereOf(issue: Issue): string {
  return issue.location.kind === "page" ? `Trang, phần ${issue.location.section}` : `Đoạn "${issue.location.excerpt}"`;
}

export function planFixes(issues: readonly Issue[], options: FixPlanOptions = {}): FixPlan {
  const includeText = options.includeText ?? true;
  const manual: Issue[] = [];
  const entries: Array<{ fix: Fix; where: string; ruleIds: string[] }> = [];

  for (const issue of issues) {
    const fix = issue.fix;
    if (!fix || (!includeText && (fix.kind === "replace" || fix.kind === "edits")) || options.skipRule?.(issue.ruleId)) {
      manual.push(issue);
      continue;
    }
    // Các cách sửa của cùng một đoạn hoặc phần: gộp nếu cùng loại và gộp được; hai cách thay chữ khác nhau trên một đoạn
    // không gộp được (cách sau dựa vào chữ chưa sửa của cách trước) nên hoãn sang lượt kiểm tra lại.
    const same = entries.find((entry) => targetOf(entry.fix) === targetOf(fix) && entry.fix.kind === fix.kind);
    if (!same) {
      entries.push({ fix, where: whereOf(issue), ruleIds: [issue.ruleId] });
      continue;
    }
    if (JSON.stringify(same.fix) === JSON.stringify(fix)) {
      same.ruleIds.push(issue.ruleId);
      continue;
    }
    const merged = merge(same.fix, fix);
    if (merged) {
      same.fix = merged;
      same.ruleIds.push(issue.ruleId);
    } else {
      manual.push(issue);
    }
  }

  const position = (fix: Fix): number => (fix.kind === "page" ? fix.section - 1_000_000 : fix.index);
  // Cách sửa thêm đoạn làm lệch chỉ số các đoạn sau nên đặt sau cùng, chỉ số lớn trước.
  const inserts = (fix: Fix): boolean => fix.kind === "insertAfter";
  entries.sort((a, b) => Number(inserts(a.fix)) - Number(inserts(b.fix)) || (inserts(a.fix) ? position(b.fix) - position(a.fix) : position(a.fix) - position(b.fix) || KIND_ORDER[a.fix.kind] - KIND_ORDER[b.fix.kind]));
  return { fixes: entries.map((e) => e.fix), changes: entries.map((e) => ({ label: e.fix.label, where: e.where, ruleIds: e.ruleIds })), manual };
}

export interface SimulatedFixAll {
  doc: CheckDocument;
  rounds: number;
  /** Lỗi còn lại sau khi sửa xong mọi lượt. */
  remaining: Issue[];
}

/** Tối đa bấy nhiêu lượt "sửa rồi kiểm tra lại"; lượt sau xử lý các lỗi mà lượt trước hoãn hoặc mới lộ ra. */
export const MAX_FIX_ROUNDS = 3;

/** Giả lập "Sửa tất cả" trên bản sao: sửa, kiểm tra lại, lặp tới khi không còn gì sửa được. */
export function simulateFixAll(doc: CheckDocument, check: (doc: CheckDocument) => Issue[], options: FixPlanOptions = {}): SimulatedFixAll {
  let current = doc;
  let rounds = 0;
  for (; rounds < MAX_FIX_ROUNDS; rounds += 1) {
    const plan = planFixes(check(current), options);
    if (plan.fixes.length === 0) break;
    current = plan.fixes.reduce((d, fix) => applyFix(d, fix), current);
  }
  return { doc: current, rounds, remaining: check(current) };
}
