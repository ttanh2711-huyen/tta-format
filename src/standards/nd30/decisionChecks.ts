import { checkBasisBlocks } from "../../checker/basisRules";
import { paragraphIssue } from "../../checker/common";
import type { CheckDocument, Issue } from "../../checker/types";
import type { StandardProfile } from "../types";

/**
 * Quy tắc phần căn cứ của chuẩn Chính quyền (T16): dòng chốt "QUYẾT ĐỊNH:", "QUYẾT NGHỊ:" in đậm; các căn cứ (trên dòng
 * chốt hoặc ngay dưới tên loại, trích yếu) theo Phụ lục I, II.6 (xem `checker/basisRules`).
 */

const VERDICT = /^QUYẾT\s+(ĐỊNH|NGHỊ)\s*:?$/;

export function checkNd30Basis(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const issues: Issue[] = [];
  for (const verdict of doc.paragraphs) {
    const t = verdict.resolved.text.trim();
    if (verdict.inTable || !VERDICT.test(t) || verdict.resolved.bold === true) continue;
    issues.push(paragraphIssue(verdict, { ruleId: "basis.verdict.bold", severity: "error", message: `Dòng "${t}" phải in đậm.`, reference: profile.elements.canCu.source, fix: { kind: "font", index: verdict.index, bold: true, label: "In đậm dòng chốt" } }));
  }
  return [...issues, ...checkBasisBlocks(doc, profile)];
}
