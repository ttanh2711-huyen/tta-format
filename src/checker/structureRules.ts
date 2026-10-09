import { paragraphIssue, replaceFix } from "./common";
import type { CheckDocument, CheckParagraph, Issue } from "./types";
import type { StandardProfile } from "../standards/types";

/**
 * Quy tắc cấu trúc Phần/Chương/Mục/Điều dùng chung (T18, TD6), nhận biết qua style của chuẩn: tên Chương/Phần và tên Mục
 * (và tên phụ lục) phải in hoa và đậm, "Điều" phải đậm; nhãn "Điều N." gõ tay phải liên tục. Nhãn tự đánh số của Word
 * không nằm trong chữ của đoạn nên chỉ kiểm tra số thứ tự khi người dùng gõ tay.
 */

const DIEU_LABEL = /^Điều\s+(\d+)\s*\./;

export interface StructureTitle {
  /** Tên style của đoạn tên (ví dụ "ND30_TieuDePhanChuong"). */
  style: string;
  /** Mã rút gọn dùng trong mã quy tắc: `structure.<rule>.upper|bold`. */
  rule: string;
  label: string;
  /** Mã thành phần của chuẩn, lấy điều khoản tham chiếu. */
  elementId: string;
}

function styled(doc: CheckDocument, name: string): CheckParagraph[] {
  return doc.paragraphs.filter((p) => !p.inTable && p.styleName === name);
}

export function checkStructureTitles(doc: CheckDocument, profile: StandardProfile, titles: readonly StructureTitle[], dieuStyle: string): Issue[] {
  const issues: Issue[] = [];
  for (const { style, rule, label, elementId } of titles) {
    const source = profile.elements[elementId]?.source ?? "";
    for (const p of styled(doc, style)) {
      const text = p.resolved.text.trim();
      if (text === "") continue;
      if (text !== text.toUpperCase()) {
        issues.push(paragraphIssue(p, { ruleId: `structure.${rule}.upper`, severity: "error", message: `${label} phải in hoa.`, reference: source, actual: text, expected: text.toUpperCase(), fix: replaceFix(p, text.toUpperCase(), `In hoa ${label.toLowerCase()}`) }));
      }
      if (p.resolved.bold !== true) {
        issues.push(paragraphIssue(p, { ruleId: `structure.${rule}.bold`, severity: "error", message: `${label} phải in đậm.`, reference: source, fix: { kind: "font", index: p.index, bold: true, label: `In đậm ${label.toLowerCase()}` } }));
      }
    }
  }
  const dieuSource = profile.elements.dieu.source;
  const dieu = styled(doc, dieuStyle);
  for (const p of dieu) {
    if (p.resolved.text.trim() !== "" && p.resolved.bold !== true) {
      issues.push(paragraphIssue(p, { ruleId: "structure.dieu.bold", severity: "error", message: "Từ \"Điều\", số thứ tự và tiêu đề của Điều phải in đậm.", reference: dieuSource, fix: { kind: "font", index: p.index, bold: true, label: "In đậm Điều" } }));
    }
  }
  let previous: number | undefined;
  for (const p of dieu) {
    const m = DIEU_LABEL.exec(p.resolved.text.trim());
    if (!m) continue;
    const n = Number(m[1]);
    if (previous !== undefined && n !== previous + 1) {
      issues.push(paragraphIssue(p, { ruleId: "structure.dieu.sequence", severity: "warning", message: `Điều ${n} đứng sau Điều ${previous}; số Điều gõ tay phải liên tục.`, reference: dieuSource, actual: `Điều ${n}`, expected: `Điều ${previous + 1}` }));
    }
    previous = n;
  }
  return issues;
}
