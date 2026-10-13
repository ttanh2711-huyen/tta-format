import { editDistance, normalizeTitle } from "../standards/detectStandard";
import type { Fix, Issue, CheckParagraph } from "./types";

/** Hàm dùng chung của các bộ quy tắc khối đầu, khối cuối: so khớp chữ chuẩn, ngày tháng, tạo lỗi và cách sửa. */

/** Tỉ lệ sai khác tối đa (so với độ dài chữ chuẩn) vẫn coi là viết sai chữ chuẩn. */
export const NEAR_RATIO = 0.15;
const DAYS = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

export function pad2(n: string): string {
  return n.length === 1 ? `0${n}` : n;
}

export function paragraphIssue(p: CheckParagraph, issue: Omit<Issue, "location">): Issue {
  const text = p.resolved.text.replace(/\s+/g, " ").trim();
  return { ...issue, location: { kind: "paragraph", index: p.index, excerpt: text.length > 40 ? `${text.slice(0, 40)}…` : text } };
}

export const replaceFix = (p: CheckParagraph, replacement: string, label: string): Fix => ({ kind: "replace", index: p.index, find: p.resolved.text.trim(), replacement, label });

/** So khớp một đoạn với chữ chuẩn: "exact" đúng từng ký tự, "near" gần đúng (sai dấu, dấu gạch, viết sai nhẹ), "no" không liên quan. */
export function matchTitle(text: string, canonical: string): "exact" | "near" | "no" {
  const trimmed = text.trim();
  if (trimmed === canonical) return "exact";
  const a = normalizeTitle(trimmed);
  const b = normalizeTitle(canonical);
  if (a === "") return "no";
  return editDistance(a, b) / b.length <= NEAR_RATIO ? "near" : "no";
}

export function checkTitleParagraph(p: CheckParagraph, canonical: string, ruleId: string, label: string, element: { min: number; max: number }, source: string): Issue[] {
  const kind = matchTitle(p.resolved.text, canonical);
  if (kind === "no") return [];
  if (kind === "near") {
    return [
      paragraphIssue(p, {
        ruleId: `${ruleId}.text`,
        severity: "error",
        message: `${label} viết sai: "${p.resolved.text.trim()}"; phải ghi "${canonical}".`,
        reference: source,
        actual: p.resolved.text.trim(),
        expected: canonical,
        fix: replaceFix(p, canonical, `Sửa ${label.toLowerCase()} cho đúng`),
      }),
    ];
  }
  const issues: Issue[] = [];
  if (p.resolved.bold !== true) {
    issues.push(
      paragraphIssue(p, {
        ruleId: `${ruleId}.bold`,
        severity: "error",
        message: `${label} phải in đậm.`,
        reference: source,
        actual: "không đậm",
        expected: "đậm",
        fix: { kind: "font", index: p.index, bold: true, label: `In đậm ${label.toLowerCase()}` },
      }),
    );
  }
  const size = p.resolved.sizePt;
  if (size !== undefined && (size < element.min || size > element.max)) {
    issues.push(
      paragraphIssue(p, {
        ruleId: `${ruleId}.size`,
        severity: "error",
        message: `${label} cỡ ${size} pt; phải từ ${element.min} đến ${element.max} pt.`,
        reference: source,
        actual: `${size} pt`,
        expected: `${element.min}–${element.max} pt`,
        fix: { kind: "font", index: p.index, sizePt: element.max, label: `Đặt cỡ chữ ${element.max} pt cho ${label.toLowerCase()}` },
      }),
    );
  }
  return issues;
}

/** Mẫu "Địa danh, ngày 05 tháng 01 năm 2020". */
export const DATE_PATTERN = /^([^,]{1,60},\s*)ngày(\s+)(\d*)(\s*)tháng(\s+)(\d*)(\s*)năm(\s*)(\d*)$/;

/** "Địa danh, ngày 05 tháng 01 năm 2020": ngày < 10 và tháng 1, 2 thêm số 0; ngày phải tồn tại. */
export function checkDate(p: CheckParagraph, source: string): Issue[] {
  const raw = p.resolved.text.trim();
  const m = DATE_PATTERN.exec(raw);
  if (!m) return [];
  const [, place, s1, day, s2, s3, month, s4, s5, year] = m;
  const issues: Issue[] = [];
  const problems: string[] = [];
  let fixedDay = day;
  let fixedMonth = month;
  if (day !== "" && day.length === 1) {
    fixedDay = pad2(day);
    problems.push("ngày nhỏ hơn 10 phải thêm số 0");
  }
  if (month !== "") {
    const n = Number(month);
    if (month.length === 1 && n < 3) {
      fixedMonth = pad2(month);
      problems.push("tháng 1 và 2 phải thêm số 0");
    } else if (month.length === 2 && month.startsWith("0") && n >= 3) {
      fixedMonth = String(n);
      problems.push("từ tháng 3 trở đi không thêm số 0");
    }
  }
  if (problems.length > 0) {
    const expected = `${place}ngày${s1}${fixedDay}${s2}tháng${s3}${fixedMonth}${s4}năm${s5}${year}`;
    issues.push(
      paragraphIssue(p, {
        ruleId: "header.date.format",
        severity: "error",
        message: `Ngày tháng "${raw}" chưa đúng quy cách: ${problems.join("; ")}.`,
        reference: source,
        actual: raw,
        expected: expected.trim(),
        fix: replaceFix(p, expected.trim(), "Sửa quy cách ngày tháng"),
      }),
    );
  }
  if (day !== "" && month !== "") {
    const d = Number(day);
    const mo = Number(month);
    if (mo < 1 || mo > 12 || d < 1 || d > DAYS[mo - 1]) {
      issues.push(
        paragraphIssue(p, {
          ruleId: "header.date.invalid",
          severity: "error",
          message: `Ngày ${d} tháng ${mo} không tồn tại.`,
          reference: source,
          actual: `ngày ${day} tháng ${month}`,
        }),
      );
    }
  }
  if (year !== "" && year.length !== 4) {
    issues.push(paragraphIssue(p, { ruleId: "header.date.year", severity: "warning", message: `Năm "${year}" phải gồm 4 chữ số.`, reference: source, actual: year, expected: "4 chữ số" }));
  }
  return issues;
}

/** Viết hoa chữ cái đầu mỗi chữ, còn lại viết thường ("NGUYỄN VĂN ĐỨC" thành "Nguyễn Văn Đức"). */
export function toTitleCase(text: string): string {
  return text
    .toLowerCase()
    .split(/(\s+)/)
    .map((part) => (part.trim() === "" ? part : part.charAt(0).toUpperCase() + part.slice(1)))
    .join("");
}
