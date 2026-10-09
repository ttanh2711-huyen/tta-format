import { normalizeTitle } from "../standards/detectStandard";
import type { StandardProfile } from "../standards/types";
import { paragraphIssue, replaceFix } from "./common";
import type { CheckDocument, Issue } from "./types";

/**
 * Nhận biết khung dấu chỉ mức độ khẩn (T21, TD8): đoạn nằm trong bảng có chữ trùng một dấu của chuẩn (bỏ dấu, không
 * phân biệt hoa thường) phải in hoa, in đậm và đúng cỡ chữ của thành phần "dấu chỉ mức độ khẩn".
 */
export function checkUrgencyMarks(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const rule = profile.elements.dauKhan;
  if (!rule) return [];
  const marks = new Set(profile.urgencyMarks.map((mark) => normalizeTitle(mark.text)));
  const issues: Issue[] = [];
  for (const p of doc.paragraphs) {
    if (!p.inTable) continue;
    const text = p.resolved.text.trim();
    if (text === "" || !marks.has(normalizeTitle(text))) continue;
    if (text !== text.toUpperCase()) {
      issues.push(paragraphIssue(p, { ruleId: "urgency.upper", severity: "error", message: "Dấu chỉ mức độ khẩn phải in hoa.", reference: rule.source, actual: text, expected: text.toUpperCase(), fix: replaceFix(p, text.toUpperCase(), "In hoa dấu khẩn") }));
    }
    if (p.resolved.bold !== true) {
      issues.push(paragraphIssue(p, { ruleId: "urgency.bold", severity: "error", message: "Dấu chỉ mức độ khẩn phải in đậm.", reference: rule.source, fix: { kind: "font", index: p.index, bold: true, label: "In đậm dấu khẩn" } }));
    }
    const size = p.resolved.sizePt;
    if (size !== undefined && (size < rule.sizePt.min || size > rule.sizePt.max)) {
      const fixed = Math.min(rule.sizePt.max, Math.max(rule.sizePt.min, size));
      issues.push(paragraphIssue(p, { ruleId: "urgency.size", severity: "error", message: `Dấu chỉ mức độ khẩn cỡ ${size} pt; phải từ ${rule.sizePt.min} đến ${rule.sizePt.max} pt.`, reference: rule.source, fix: { kind: "font", index: p.index, sizePt: fixed, label: `Đặt cỡ chữ ${fixed} pt` } }));
    }
  }
  return issues;
}
