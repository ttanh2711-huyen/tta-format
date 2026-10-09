import { runChecks } from "../checker/engine";
import { MAX_FIX_ROUNDS, planFixes, type FixPlanOptions } from "../checker/fixAll";
import { headerFooterRules } from "../checker/ruleSets";
import type { Issue } from "../checker/types";
import type { StandardProfile } from "../standards/types";
import { applyFixes } from "./applyFix";
import { readCheckDocument } from "./readDocument";

export interface FixAllRun {
  applied: number;
  skipped: string[];
  rounds: number;
  /** Lỗi còn lại sau lượt cuối, đọc lại từ tài liệu. */
  remaining: Issue[];
}

/** Đọc tài liệu và kiểm tra theo chuẩn, gồm cả quy tắc khối đầu, khối cuối. */
export async function checkDocument(profile: StandardProfile): Promise<Issue[]> {
  return runChecks(await readCheckDocument(), profile, headerFooterRules(profile));
}

/**
 * Sửa, đọc lại, sửa tiếp tới khi không còn gì sửa được (tối đa `MAX_FIX_ROUNDS` lượt): lượt sau xử lý các lỗi
 * mà lượt trước hoãn (hai cách thay chữ trên một đoạn) hoặc mới lộ ra. Dừng sớm nếu một lượt không sửa được gì.
 */
export async function runFixAll(profile: StandardProfile, options: FixPlanOptions = {}): Promise<FixAllRun> {
  let applied = 0;
  let rounds = 0;
  const skipped: string[] = [];
  let issues = await checkDocument(profile);
  for (; rounds < MAX_FIX_ROUNDS; rounds += 1) {
    const plan = planFixes(issues, options);
    if (plan.fixes.length === 0) break;
    const result = await applyFixes(plan.fixes);
    applied += result.applied;
    skipped.push(...result.skipped);
    issues = await checkDocument(profile);
    if (result.applied === 0) break;
  }
  return { applied, skipped, rounds, remaining: issues };
}
