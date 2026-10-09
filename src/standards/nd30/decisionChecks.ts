import { paragraphIssue } from "../../checker/common";
import type { CheckDocument, CheckParagraph, Fix, Issue } from "../../checker/types";
import type { StandardProfile } from "../types";

/**
 * Quy tắc cho phần căn cứ của nghị quyết, quyết định (T16): khối các dòng "Căn cứ ..." liền ngay trên dòng "QUYẾT ĐỊNH:"
 * hoặc "QUYẾT NGHỊ:". Mỗi căn cứ in nghiêng, cỡ 13 đến 14, cuối dòng ";" và căn cứ cuối kết thúc bằng ","; dòng chốt in đậm.
 */

const VERDICT = /^QUYẾT\s+(ĐỊNH|NGHỊ)\s*:?$/;
/** Không dùng `\b`: trong JavaScript chữ có dấu không được coi là ký tự chữ nên ranh giới từ sai. */
const BASIS = /^(Căn cứ|Xét đề nghị|Theo đề nghị|Xét|Theo)(?=\s|$)/;
/** Chữ cuối đoạn dùng để thay dấu cuối dòng mà không phải tìm cả đoạn dài (Word chỉ tìm được tối đa 255 ký tự). */
const TAIL_CHARS = 30;

function endingFix(p: CheckParagraph, ending: string, label: string): Fix {
  const text = p.resolved.text.trim();
  const tail = text.slice(-TAIL_CHARS);
  return { kind: "replace", index: p.index, find: tail, replacement: `${tail.replace(/[;.,]+$/, "")}${ending}`, label };
}

export function checkNd30Basis(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const issues: Issue[] = [];
  const canCu = profile.elements.canCu;
  const body = doc.paragraphs.filter((p) => !p.inTable && p.resolved.text.trim() !== "");
  body.forEach((verdict, at) => {
    if (!VERDICT.test(verdict.resolved.text.trim())) return;
    if (verdict.resolved.bold !== true) {
      issues.push(paragraphIssue(verdict, { ruleId: "basis.verdict.bold", severity: "error", message: `Dòng "${verdict.resolved.text.trim()}" phải in đậm.`, reference: canCu.source, fix: { kind: "font", index: verdict.index, bold: true, label: "In đậm dòng chốt" } }));
    }
    const block: CheckParagraph[] = [];
    for (let i = at - 1; i >= 0 && BASIS.test(body[i].resolved.text.trim()); i -= 1) block.unshift(body[i]);
    block.forEach((p, i) => {
      const last = i === block.length - 1;
      const text = p.resolved.text.trim();
      if (p.resolved.italic !== true) {
        issues.push(paragraphIssue(p, { ruleId: "basis.italic", severity: "error", message: "Căn cứ ban hành phải in nghiêng.", reference: canCu.source, fix: { kind: "font", index: p.index, italic: true, label: "In nghiêng căn cứ" } }));
      }
      const size = p.resolved.sizePt;
      if (size !== undefined && (size < canCu.sizePt.min || size > canCu.sizePt.max)) {
        const fixed = Math.min(canCu.sizePt.max, Math.max(canCu.sizePt.min, size));
        issues.push(paragraphIssue(p, { ruleId: "basis.size", severity: "error", message: `Căn cứ cỡ ${size} pt; phải từ ${canCu.sizePt.min} đến ${canCu.sizePt.max} pt.`, reference: canCu.source, fix: { kind: "font", index: p.index, sizePt: fixed, label: `Đặt cỡ chữ ${fixed} pt` } }));
      }
      const wanted = last ? "," : ";";
      if (!text.endsWith(wanted)) {
        issues.push(
          paragraphIssue(p, {
            ruleId: "basis.punctuation",
            severity: "error",
            message: last ? "Căn cứ cuối cùng kết thúc bằng dấu phẩy." : "Mỗi căn cứ kết thúc bằng dấu chấm phẩy.",
            reference: canCu.source,
            actual: text.slice(-1),
            expected: wanted,
            fix: endingFix(p, wanted, last ? "Kết thúc căn cứ cuối bằng dấu phẩy" : "Kết thúc căn cứ bằng dấu chấm phẩy"),
          }),
        );
      }
    });
  });
  return issues;
}
