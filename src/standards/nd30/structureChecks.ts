import { paragraphIssue, replaceFix } from "../../checker/common";
import type { CheckDocument, CheckParagraph, Issue } from "../../checker/types";
import type { StandardProfile } from "../types";

/**
 * Quy tắc cấu trúc Phần/Chương/Mục/Điều (T18), nhận biết qua style của chuẩn: tên Chương và tên Mục phải in hoa và đậm,
 * "Điều" phải đậm; nhãn "Điều N." gõ tay phải liên tục. Nhãn tự đánh số của Word không nằm trong chữ của đoạn nên
 * chỉ kiểm tra số thứ tự khi người dùng gõ tay.
 */

const DIEU_LABEL = /^Điều\s+(\d+)\s*\./;

function styled(doc: CheckDocument, name: string): CheckParagraph[] {
  return doc.paragraphs.filter((p) => !p.inTable && p.styleName === name);
}

export function checkNd30Structure(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const issues: Issue[] = [];
  const prefix = profile.styles.prefix;
  const titles: Array<{ style: string; rule: "chuong" | "muc" | "phuLuc"; label: string; elementId: string }> = [
    { style: `${prefix}TieuDePhanChuong`, rule: "chuong", label: "Tên Chương", elementId: "tieuDePhanChuong" },
    { style: `${prefix}TieuDeMucTieuMuc`, rule: "muc", label: "Tên Mục", elementId: "tieuDeMucTieuMuc" },
    { style: `${prefix}PhuLucTieuDe`, rule: "phuLuc", label: "Tên phụ lục", elementId: "phuLucTieuDe" },
  ];
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
  const dieu = styled(doc, `${prefix}Dieu`);
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
