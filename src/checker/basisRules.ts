import type { StandardProfile } from "../standards/types";
import { paragraphIssue } from "./common";
import type { CheckDocument, CheckParagraph, Fix, Issue } from "./types";

/**
 * Căn cứ ban hành văn bản, cho mọi loại văn bản có phần căn cứ, cả hai chuẩn:
 * - Nghị định 30 (Phụ lục I, II.6): chữ in thường, nghiêng, cỡ 13 đến 14, dưới tên loại và trích yếu; mỗi căn cứ một
 *   dòng, cuối dòng ";", dòng cuối "."; không có gạch đầu dòng.
 * - Hướng dẫn 05 (6.2.3 và 6.2.2): dưới tên loại và trích yếu; trước mỗi căn cứ một dấu gạch ngang "-", sau mỗi căn cứ
 *   ";", căn cứ cuối "."; chữ đứng như phần nội dung.
 * Khối căn cứ là các dòng "Căn cứ…", "Xét đề nghị…", "Theo đề nghị…" liền nhau, nằm ngay trên dòng chốt ("QUYẾT ĐỊNH:",
 * "QUYẾT NGHỊ:"…) hoặc ngay dưới tên loại, trích yếu (công văn không có phần căn cứ nên không xét).
 */

/** Không dùng `\b`: chữ có dấu không được coi là ký tự chữ trong JavaScript. */
export const BASIS_LINE = /^[-–—]?\s*(Căn cứ|Xét đề nghị|Theo đề nghị|Xét|Theo)(?=\s|$)/u;
const BASIS_FIRST = /^[-–—]?\s*Căn cứ(?=\s|$)/u;
const VERDICT = /^(QUYẾT\s+(ĐỊNH|NGHỊ)|KẾT\s+LUẬN|CHỈ\s+THỊ|QUY\s+ĐỊNH)\s*:\s*$/u;
const KINH = /^(Kính\s+(gửi|trình)|[-–—+•]\s)/u;
const DASH = /^[-–—]\s*/u;
/** Chữ đầu, cuối đoạn dùng để thay mà không phải tìm cả đoạn dài (Word chỉ tìm được tối đa 255 ký tự). */
const EDGE_CHARS = 30;

const text = (p: CheckParagraph): string => p.resolved.text.trim();
const isUpper = (s: string): boolean => {
  const letters = s.replace(/[^\p{L}]/gu, "");
  return letters.length > 0 && letters === letters.toUpperCase();
};

/** Các khối căn cứ của văn bản (mỗi khối là các đoạn liền nhau), không trùng nhau. */
export function findBasisBlocks(doc: CheckDocument): CheckParagraph[][] {
  const body = doc.paragraphs.filter((p) => !p.inTable && text(p) !== "");
  const blocks: CheckParagraph[][] = [];
  const taken = new Set<number>();
  const push = (block: CheckParagraph[]): void => {
    // Dòng dẫn kết thúc bằng ":" ("Căn cứ tình hình thực tế, ... như sau:") là câu văn thân bài, không phải căn cứ.
    const trimmed = text(block[block.length - 1]).endsWith(":") ? block.slice(0, -1) : block;
    if (trimmed.length === 0 || trimmed.some((p) => taken.has(p.index))) return;
    trimmed.forEach((p) => taken.add(p.index));
    blocks.push(trimmed);
  };

  // 1. Ngay trên dòng chốt của nghị quyết, quyết định.
  body.forEach((p, at) => {
    if (!VERDICT.test(text(p))) return;
    const block: CheckParagraph[] = [];
    for (let i = at - 1; i >= 0 && BASIS_LINE.test(text(body[i])); i -= 1) block.unshift(body[i]);
    if (block.length > 0) push(block);
  });

  // 2. Ngay dưới tên loại, trích yếu (và Kính gửi, Kính trình nếu có): trước khối chỉ có đoạn căn giữa và dòng Kính gửi.
  let sawTitle = false;
  for (let i = 0; i < body.length; i += 1) {
    const p = body[i];
    const t = text(p);
    if (BASIS_FIRST.test(t)) {
      if (!sawTitle) break;
      const block: CheckParagraph[] = [];
      for (let j = i; j < body.length && BASIS_LINE.test(text(body[j])); j += 1) block.push(body[j]);
      push(block);
      break;
    }
    const centered = p.resolved.align === "center";
    if (centered && isUpper(t)) sawTitle = true;
    if (!centered && !KINH.test(t)) break;
  }
  return blocks;
}

export interface BasisRule {
  /** Có gạch đầu dòng "-" trước mỗi căn cứ (Đảng) hay không (Chính quyền). */
  dash: boolean;
  italic: boolean;
  sizePt: { min: number; max: number };
  source: string;
}

/** Quy tắc căn cứ theo chuẩn. */
export function basisRuleOf(profile: StandardProfile): BasisRule {
  if (profile.headerKind === "dang") {
    return { dash: true, italic: false, sizePt: profile.elements.noiDung.sizePt, source: "Hướng dẫn 05-HD/VPTW, mục 6.2.3 (căn cứ ban hành) và 6.2.2 (nội dung)" };
  }
  const canCu = profile.elements.canCu;
  return { dash: false, italic: true, sizePt: canCu.sizePt, source: `Phụ lục I, ${canCu.source}` };
}

function endingFix(p: CheckParagraph, ending: string, label: string): Fix {
  const tail = text(p).slice(-EDGE_CHARS);
  return { kind: "replace", index: p.index, find: tail, replacement: `${tail.replace(/[;.,:]+$/u, "")}${ending}`, label };
}

function dashFix(p: CheckParagraph, dash: boolean): Fix {
  const head = text(p).slice(0, EDGE_CHARS);
  const bare = head.replace(DASH, "");
  return { kind: "replace", index: p.index, find: head, replacement: dash ? `- ${bare}` : bare, label: dash ? "Thêm gạch đầu dòng trước căn cứ" : "Bỏ gạch đầu dòng trước căn cứ" };
}

/** Kiểm tra từng khối căn cứ: gạch đầu dòng, kiểu chữ, cỡ chữ, dấu cuối mỗi căn cứ. */
export function checkBasisBlocks(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const rule = basisRuleOf(profile);
  const issues: Issue[] = [];
  for (const block of findBasisBlocks(doc)) {
    block.forEach((p, i) => {
      const t = text(p);
      const last = i === block.length - 1;
      if (DASH.test(t) !== rule.dash) {
        issues.push(paragraphIssue(p, { ruleId: "basis.dash", severity: "error", message: rule.dash ? "Trước mỗi căn cứ có một dấu gạch ngang (-)." : "Căn cứ ban hành không có gạch đầu dòng.", reference: rule.source, fix: dashFix(p, rule.dash) }));
      }
      if (rule.italic && p.resolved.italic !== true) {
        issues.push(paragraphIssue(p, { ruleId: "basis.italic", severity: "error", message: "Căn cứ ban hành phải in nghiêng.", reference: rule.source, fix: { kind: "font", index: p.index, italic: true, label: "In nghiêng căn cứ" } }));
      }
      if (!rule.italic && p.resolved.italic === true) {
        issues.push(paragraphIssue(p, { ruleId: "basis.upright", severity: "error", message: "Căn cứ ban hành trình bày chữ đứng như nội dung văn bản, không in nghiêng.", reference: rule.source, fix: { kind: "font", index: p.index, italic: false, label: "Bỏ in nghiêng căn cứ" } }));
      }
      const size = p.resolved.sizePt;
      if (size !== undefined && (size < rule.sizePt.min || size > rule.sizePt.max)) {
        const fixed = Math.min(rule.sizePt.max, Math.max(rule.sizePt.min, size));
        issues.push(paragraphIssue(p, { ruleId: "basis.size", severity: "error", message: `Căn cứ cỡ ${size} pt; phải từ ${rule.sizePt.min} đến ${rule.sizePt.max} pt.`, reference: rule.source, fix: { kind: "font", index: p.index, sizePt: fixed, label: `Đặt cỡ chữ ${fixed} pt` } }));
      }
      const wanted = last ? "." : ";";
      if (!t.endsWith(wanted) || /[;.,:]{2,}$/u.test(t)) {
        issues.push(
          paragraphIssue(p, {
            ruleId: "basis.punctuation",
            severity: "error",
            message: last ? "Căn cứ cuối cùng kết thúc bằng dấu chấm (.)." : "Mỗi căn cứ kết thúc bằng dấu chấm phẩy (;).",
            reference: rule.source,
            actual: t.slice(-1),
            expected: wanted,
            fix: endingFix(p, wanted, last ? "Kết thúc căn cứ cuối bằng dấu chấm" : "Kết thúc căn cứ bằng dấu chấm phẩy"),
          }),
        );
      }
    });
  }
  return issues;
}
