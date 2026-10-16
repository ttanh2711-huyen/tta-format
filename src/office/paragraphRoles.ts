import type { StandardProfile } from "../standards/types";

/**
 * Vai trò của đoạn thân văn bản (issue #3), nhận diện theo ký hiệu đầu đoạn để "Định dạng nội dung" áp đúng style đề
 * mục thay vì áp style thân bài cho mọi đoạn:
 * - chuong: "Chương I", "Phần thứ nhất", "PHẦN I"; tenChuong: dòng in hoa ngay sau;
 * - muc: "Mục 1", "Mục II"; tenMuc: dòng in hoa ngay sau;
 * - dieu: "Điều 5."; khoan: "1. " (theo sau là chữ); diem: "a) ", "đ) ", "1.1. ";
 * - deMuc: đề mục số La Mã "I. ", "II. "; y: ý gạch đầu dòng "- ", "+ ";
 * - noiDung: còn lại.
 * Quy tắc đòi ký hiệu đứng đầu đoạn và theo sau đúng kiểu, để "Chương trình…", "Mục tiêu…", "1.500 hộ…" vẫn là thân bài.
 */
export type ParagraphRole = "chuong" | "tenChuong" | "muc" | "tenMuc" | "dieu" | "deMuc" | "khoan" | "diem" | "y" | "noiDung";

const ROMAN = "[IVXLC]+";
const CHAPTER = new RegExp(`^(?:(?:Chương|CHƯƠNG)\\s+(?:${ROMAN}|\\d+)(?![\\p{L}\\d])|(?:Phần|PHẦN)\\s+(?:thứ\\s+\\p{L}+|THỨ\\s+\\p{L}+|${ROMAN}|\\d+)(?![\\p{L}\\d]))`, "u");
const SECTION = new RegExp(`^(?:Mục|MỤC)\\s+(?:${ROMAN}|\\d+)(?![\\p{L}\\d])`, "u");
const ARTICLE = /^(?:Điều|ĐIỀU)\s+\d+[a-zđ]?\./u;
const ROMAN_HEADING = new RegExp(`^${ROMAN}\\.\\s+\\p{L}`, "u");
// "1.1. Về…" hoặc "1.1 Về…" (chữ hoa ngay sau); không nhầm số có dấu chấm ngăn cách hàng nghìn ("1.500 hộ").
const SUB_NUMBER = /^\d{1,2}\.\d{1,2}(?:\.\s+\p{L}|\s+\p{Lu})/u;
const CLAUSE = /^\d{1,2}\.\s+\p{L}/u;
const POINT = /^[a-zđ]\)\s/u;
const BULLET = /^[-+–•]\s/u;

/** Vai trò của một đoạn, chỉ xét chữ của chính đoạn đó (chưa biết đoạn đứng trước). */
export function detectRole(text: string): ParagraphRole {
  const t = text.trim();
  if (CHAPTER.test(t)) return "chuong";
  if (SECTION.test(t)) return "muc";
  if (ARTICLE.test(t)) return "dieu";
  if (ROMAN_HEADING.test(t)) return "deMuc";
  if (SUB_NUMBER.test(t)) return "diem";
  if (CLAUSE.test(t)) return "khoan";
  if (POINT.test(t)) return "diem";
  if (BULLET.test(t)) return "y";
  return "noiDung";
}

const isUpperLine = (text: string): boolean => {
  const letters = text.replace(/[^\p{L}]/gu, "");
  return letters.length > 0 && letters === letters.toUpperCase();
};

/** Vai trò cho cả dãy đoạn: dòng in hoa ngay sau "Chương"/"Phần" là tên chương, sau "Mục" là tên mục. */
export function assignRoles(texts: readonly string[]): ParagraphRole[] {
  const roles: ParagraphRole[] = [];
  texts.forEach((text, i) => {
    const own = detectRole(text);
    const previous = i > 0 ? roles[i - 1] : undefined;
    if (own === "noiDung" && isUpperLine(text) && previous === "chuong") roles.push("tenChuong");
    else if (own === "noiDung" && isUpperLine(text) && previous === "muc") roles.push("tenMuc");
    else roles.push(own);
  });
  return roles;
}

/** Thành phần thể thức (style) của từng vai trò, theo thứ tự ưu tiên; chuẩn nào có thành phần nào thì dùng thành phần đó. */
const ELEMENTS: Record<ParagraphRole, readonly string[]> = {
  chuong: ["phanChuong"],
  tenChuong: ["tieuDePhanChuong", "tenPhanChuong"],
  muc: ["mucTieuMuc", "mucSo"],
  tenMuc: ["tieuDeMucTieuMuc", "tenMuc"],
  dieu: ["dieu"],
  khoan: ["khoan"],
  diem: ["diem"],
  deMuc: [],
  y: [],
  noiDung: [],
};

const pascal = (id: string): string => id.charAt(0).toUpperCase() + id.slice(1);

/** Style áp cho vai trò; vai trò không có style riêng dùng style nội dung (đề mục La Mã thêm in đậm). */
export function roleStyle(role: ParagraphRole, profile: StandardProfile): { styleName: string; bold?: true } {
  const planned = new Set(profile.styles.elements.map((item) => item.elementId));
  const element = ELEMENTS[role].find((id) => planned.has(id));
  if (element) return { styleName: `${profile.styles.prefix}${pascal(element)}` };
  return role === "deMuc" ? { styleName: `${profile.styles.prefix}NoiDung`, bold: true } : { styleName: `${profile.styles.prefix}NoiDung` };
}

/** Vai trò căn giữa (tên chương, mục): không ghi giãn dòng Exactly như đoạn văn. */
export const CENTERED_ROLES: ReadonlySet<ParagraphRole> = new Set(["chuong", "tenChuong", "muc", "tenMuc"]);
