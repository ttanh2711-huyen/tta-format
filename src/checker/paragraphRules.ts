import { isFirstLineIndentValid, isFontSizeValid, isLineSpacingValid } from "../standards/validate";
import type { ElementRule, StandardProfile } from "../standards/types";
import type { CheckDocument, CheckParagraph, Fix, Issue } from "./types";

/** Đoạn có ít nhất bấy nhiêu ký tự mới coi là văn xuôi để kiểm tra căn lề, thụt dòng, giãn đoạn, giãn dòng. */
export const PROSE_MIN_CHARS = 40;
const EXCERPT_CHARS = 40;
const SPACING_TOLERANCE = 0.01;

function vn(n: number): string {
  return String(Math.round(n * 100) / 100).replace(".", ",");
}

function excerpt(text: string): string {
  const clean = text.replace(/\s+/g, " ").trim();
  return clean.length > EXCERPT_CHARS ? `${clean.slice(0, EXCERPT_CHARS)}…` : clean;
}

/** Thành phần thể thức mà style của đoạn đại diện, theo tiền tố style của chuẩn (ví dụ "ND30_Dieu" là "dieu"). */
function elementOf(styleName: string | undefined, profile: StandardProfile): { id: string; rule: ElementRule } | undefined {
  const prefix = profile.styles.prefix;
  if (!styleName || !styleName.startsWith(prefix)) return undefined;
  const suffix = styleName.slice(prefix.length);
  const id = suffix.charAt(0).toLowerCase() + suffix.slice(1);
  const rule = profile.elements[id];
  return rule ? { id, rule } : undefined;
}

function isBodyLike(p: CheckParagraph, element: { id: string } | undefined): boolean {
  const r = p.resolved;
  const alignOk = r.align === undefined || r.align === "left" || r.align === "both";
  const indented = (r.hangingIndentMm ?? 0) > 0 || (r.leftIndentMm ?? 0) > 0;
  return (element === undefined || element.id === "noiDung") && alignOk && !indented;
}

function describeLine(r: CheckParagraph["resolved"]): string {
  if (r.lineExactPt !== undefined) return `Exactly ${vn(r.lineExactPt)} pt`;
  if (r.lineAtLeastPt !== undefined) return `tối thiểu ${vn(r.lineAtLeastPt)} pt`;
  return `bội số ${vn(r.lineMultiple ?? 1)} dòng`;
}

/** Giãn dòng đưa về giá trị mặc định của chuẩn, kẹp vào khoảng cho phép; đổi kiểu giãn dòng nếu chuẩn quy định kiểu khác. */
function lineFix(index: number, profile: StandardProfile): Fix {
  const rule = profile.content.lineSpacing;
  const wanted = profile.styles.defaults.lineSpacing;
  if (rule.kind === "multiple") {
    const lines = wanted.kind === "multiple" ? Math.min(rule.max, Math.max(rule.min, wanted.lines)) : rule.min;
    return { kind: "paragraph", index, lineMultiple: lines, label: `Giãn dòng ${vn(lines)} dòng` };
  }
  const pt = wanted.kind === "exactly" ? Math.min(rule.maxPt, Math.max(rule.minPt, wanted.pt)) : rule.minPt;
  return { kind: "exact", index, pt, label: `Giãn dòng Exactly ${vn(pt)} pt` };
}

/** Kiểm tra phông, cỡ chữ, căn lề, thụt dòng đầu, giãn đoạn và giãn dòng của các đoạn ngoài bảng. */
export function checkParagraphs(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const issues: Issue[] = [];
  const content = profile.content;
  const bodyRule = profile.elements.noiDung;
  const defaults = profile.styles.defaults;

  for (const p of doc.paragraphs) {
    const r = p.resolved;
    if (p.inTable || r.text.trim() === "") continue;
    const element = elementOf(p.styleName, profile);
    const bodyLike = isBodyLike(p, element);
    if (!bodyLike && element === undefined) continue;
    const location = { kind: "paragraph", index: p.index, excerpt: excerpt(r.text) } as const;
    const add = (issue: Omit<Issue, "location" | "severity"> & { severity?: Issue["severity"] }) => issues.push({ severity: "error", location, ...issue });

    if (r.fontName !== undefined && r.fontName.toLowerCase() !== profile.font.name.toLowerCase()) {
      add({
        ruleId: "font.name",
        message: `Phông chữ "${r.fontName}" không đúng; chuẩn ${profile.name} yêu cầu ${profile.font.name}.`,
        reference: profile.font.source,
        actual: r.fontName,
        expected: profile.font.name,
        fix: { kind: "font", index: p.index, name: profile.font.name, label: `Đổi phông chữ thành ${profile.font.name}` },
      });
    }

    const rule = element?.rule ?? bodyRule;
    if (r.sizePt !== undefined && !isFontSizeValid(profile, element?.id ?? "noiDung", r.sizePt)) {
      const range = rule.sizePt.min === rule.sizePt.max ? `${vn(rule.sizePt.min)} pt` : `${vn(rule.sizePt.min)}–${vn(rule.sizePt.max)} pt`;
      add({
        ruleId: "font.size",
        message: `Cỡ chữ ${vn(r.sizePt)} pt ngoài quy định cho ${rule.label.toLowerCase()}: ${range}.`,
        reference: rule.source,
        actual: `${vn(r.sizePt)} pt`,
        expected: range,
        fix: { kind: "font", index: p.index, sizePt: Math.min(rule.sizePt.max, Math.max(rule.sizePt.min, r.sizePt)), label: `Đặt cỡ chữ ${vn(Math.min(rule.sizePt.max, Math.max(rule.sizePt.min, r.sizePt)))} pt` },
      });
    }

    const prose = bodyLike && (r.text.trim().length >= PROSE_MIN_CHARS || element?.id === "noiDung");
    if (!prose) continue;

    if (content.align === "justify" && r.align !== "both") {
      add({
        ruleId: "paragraph.align",
        message: "Đoạn nội dung phải căn đều hai bên.",
        reference: content.source,
        actual: "căn trái",
        expected: "đều hai bên",
        fix: { kind: "paragraph", index: p.index, align: "both", label: "Căn đều hai bên" },
      });
    }

    const indent = r.firstLineIndentMm ?? 0;
    const indentTarget = content.firstLineIndentMm.includes(defaults.firstLineIndentMm) ? defaults.firstLineIndentMm : content.firstLineIndentMm[0];
    if (!isFirstLineIndentValid(profile, indent)) {
      const allowed = content.firstLineIndentMm.map((mm) => `${vn(mm)} mm`).join(" hoặc ");
      add({
        ruleId: "paragraph.firstLineIndent",
        message: indent === 0 ? `Dòng đầu của đoạn chưa thụt vào; cần ${allowed}.` : `Thụt dòng đầu ${vn(indent)} mm không đúng; cần ${allowed}.`,
        reference: content.source,
        actual: `${vn(indent)} mm`,
        expected: allowed,
        fix: { kind: "paragraph", index: p.index, firstLineIndentMm: indentTarget, label: `Thụt dòng đầu ${vn(indentTarget)} mm` },
      });
    }

    const spacing = (r.spaceBeforePt ?? 0) + (r.spaceAfterPt ?? 0);
    const spaceTarget =
      defaults.spaceBeforePt + defaults.spaceAfterPt >= content.paragraphSpacingMinPt
        ? { before: defaults.spaceBeforePt, after: defaults.spaceAfterPt }
        : { before: content.paragraphSpacingMinPt / 2, after: content.paragraphSpacingMinPt / 2 };
    if (spacing < content.paragraphSpacingMinPt - SPACING_TOLERANCE) {
      add({
        ruleId: "paragraph.spacing",
        message: `Giãn đoạn trước và sau tổng ${vn(spacing)} pt; tối thiểu ${vn(content.paragraphSpacingMinPt)} pt.`,
        reference: content.source,
        actual: `${vn(spacing)} pt`,
        expected: `tối thiểu ${vn(content.paragraphSpacingMinPt)} pt`,
        fix: { kind: "paragraph", index: p.index, spaceBeforePt: spaceTarget.before, spaceAfterPt: spaceTarget.after, label: `Giãn đoạn trước ${vn(spaceTarget.before)} pt, sau ${vn(spaceTarget.after)} pt` },
      });
    }

    const lineRule = content.lineSpacing;
    const value =
      r.lineExactPt !== undefined
        ? ({ kind: "exactly", pt: r.lineExactPt } as const)
        : r.lineAtLeastPt !== undefined
          ? undefined
          : ({ kind: "multiple", lines: r.lineMultiple ?? 1 } as const);
    if (value === undefined || !isLineSpacingValid(profile, value)) {
      const expected = lineRule.kind === "multiple" ? `bội số ${vn(lineRule.min)}–${vn(lineRule.max)} dòng` : `Exactly ${vn(lineRule.minPt)}–${vn(lineRule.maxPt)} pt`;
      add({
        ruleId: "paragraph.lineSpacing",
        message: `Giãn dòng ${describeLine(r)} không đúng; chuẩn ${profile.name} yêu cầu ${expected}.`,
        reference: content.source,
        actual: describeLine(r),
        expected,
        fix: lineFix(p.index, profile),
      });
    }

    // Chỉ báo đoạn kết thúc bằng chữ hoặc số; kết thúc bằng ":", ";", "," là đoạn dẫn vào danh sách hoặc ý tiếp theo.
    // Đề mục (in đậm cả đoạn, hoặc mở đầu bằng số La Mã "I. ") không phải câu văn nên không cần dấu chấm.
    const text = r.text.trim();
    const heading = r.bold === true || /^[IVXLC]+\.\s/.test(text);
    if (content.endsWithPeriod && !heading && /[\p{L}\p{N}]$/u.test(text)) {
      const tail = text.slice(-30);
      add({
        ruleId: "paragraph.endsWithPeriod",
        severity: "warning",
        message: "Nội dung văn bản phải kết thúc bằng dấu chấm.",
        reference: content.source,
        actual: text.slice(-1),
        expected: ".",
        fix: { kind: "replace", index: p.index, find: tail, replacement: `${tail}.`, label: "Thêm dấu chấm cuối đoạn" },
      });
    }
  }
  return issues;
}
