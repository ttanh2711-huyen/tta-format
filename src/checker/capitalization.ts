import type { StandardProfile } from "../standards/types";
import { paragraphIssue } from "./common";
import type { CheckDocument, CheckParagraph, Issue } from "./types";

/**
 * Kiểm tra viết hoa (Phụ lục II Nghị định 30/2020/NĐ-CP): đầu câu, sau dấu kết câu, tên cơ quan, tổ chức, chức vụ cố
 * định, tên loại văn bản khi viện dẫn và (chỉ chuẩn Chính quyền) phần, chương, mục, điều, khoản, điểm khi viện dẫn.
 * Đây là quy tắc dựa trên mẫu chữ nên mọi lỗi đều là cảnh báo; chỉ xét đoạn văn ngoài bảng, không in hoa toàn bộ.
 */

const LOWER = /^[\p{Ll}]$/u;
const isLower = (ch: string): boolean => LOWER.test(ch) && ch !== ch.toUpperCase();
const capitalize = (word: string): string => word.charAt(0).toUpperCase() + word.slice(1);

/** Nhãn đầu dòng, số thứ tự, dấu mở ngoặc hoặc mở nháy đứng trước chữ đầu đoạn. */
const LEADING_LABEL = /^\s*(?:(?:\d+(?:\.\d+)*[.)]|[a-zđ]\)|[IVXLC]+[.)]|[-+*•–—])\s+)?[("“'‘[]*/u;

/** Chữ viết tắt thường đứng trước dấu chấm mà không kết câu. */
const ABBREVIATIONS = new Set(["tp", "tx", "tt", "q", "p", "h", "x", "ths", "th.s", "ts", "tskh", "pgs", "gs", "ks", "bs", "cn", "ds", "kt", "tl", "tuq", "tm", "st", "vd", "tr", "v.v", "sđd", "nxb", "đ/c", "ttg"]);

/** Tên cơ quan, tổ chức, chức vụ có cách viết hoa cố định (theo ví dụ của Phụ lục II). */
const FIXED_NAMES = [
  "Ủy ban nhân dân",
  "Hội đồng nhân dân",
  "Thủ tướng Chính phủ",
  "Văn phòng Chính phủ",
  "Chủ tịch nước",
  "Quốc hội",
  "Chính phủ",
  "Ủy ban Mặt trận Tổ quốc Việt Nam",
  "Mặt trận Tổ quốc Việt Nam",
  "Đảng Cộng sản Việt Nam",
  "Ban Chấp hành Trung ương",
  "Bộ Chính trị",
  "Ban Bí thư",
  "Tòa án nhân dân",
  "Viện kiểm sát nhân dân",
  "Đoàn Thanh niên Cộng sản Hồ Chí Minh",
  "Hội Liên hiệp Phụ nữ Việt Nam",
  "Hội Cựu chiến binh Việt Nam",
  "Hội Nông dân Việt Nam",
];

/** Tên loại văn bản viết hoa chữ đầu khi là tên của một văn bản cụ thể (đi với "số" hoặc tên riêng). */
const DOC_KIND = /(?<![\p{L}])(bộ luật|(?<![Bb]ộ\s)luật|pháp lệnh)(?=\s+[\p{Lu}])|(?<![\p{L}])(nghị định|thông tư|quyết định|nghị quyết|chỉ thị|kết luận|quy định|hướng dẫn|công văn)(?=\s+số\s)/gu;

/** Phần, chương, mục, điều, khoản, điểm khi viện dẫn (chuẩn Chính quyền). Khoản, điểm chỉ khi đi tiếp tới điều, khoản. */
const REFERENCE = /(?<![\p{L}])(phần|chương|mục|điều)(?=\s+(?:[0-9]+|[IVXLC]+)(?![\p{L}]))|(?<![\p{L}])(khoản)(?=\s+[0-9]+\s+(?:điều|Điều)\s)|(?<![\p{L}])(điểm)(?=\s+[a-zđ]\s+(?:khoản|Khoản|điều|Điều)\s)/gu;

const fold = (s: string): string => s.normalize("NFC").replace(/Uỷ|uỷ/g, (m) => (m[0] === "U" ? "Ủy" : "ủy"));

interface Finding {
  ruleId: string;
  message: string;
  find: string;
  replacement: string;
}

function referenceOf(profile: StandardProfile): string {
  return profile.headerKind === "quocHieu" ? "Phụ lục II Nghị định 30/2020/NĐ-CP (viết hoa)" : "Quy tắc viết hoa thông dụng theo Phụ lục II Nghị định 30, áp dụng cho văn bản Đảng (cần đối chiếu Hướng dẫn 05)";
}

function sentenceStart(text: string): Finding | undefined {
  const label = LEADING_LABEL.exec(text)?.[0] ?? "";
  const rest = text.slice(label.length);
  const word = /^[\p{L}]+/u.exec(rest)?.[0];
  if (!word || !isLower(word[0])) return undefined;
  return { ruleId: "cap.sentenceStart", message: `Đầu câu phải viết hoa: "${word}".`, find: `${label}${word}`, replacement: `${label}${capitalize(word)}` };
}

function afterPeriod(text: string): Finding[] {
  const out: Finding[] = [];
  for (const m of text.matchAll(/([\p{L}\p{N}./]*)([.!?])(\s+)([\p{L}]+)/gu)) {
    const [, before, mark, space, word] = m;
    if (!isLower(word[0])) continue;
    if (mark === "." && (ABBREVIATIONS.has(before.toLowerCase()) || /^\p{L}$/u.test(before) || /\.$/.test(before) || before === "")) continue;
    if (mark === "." && /\.\./.test(text.slice(Math.max(0, (m.index ?? 0) + before.length - 2), (m.index ?? 0) + before.length + 1))) continue;
    out.push({ ruleId: "cap.afterPeriod", message: `Sau dấu "${mark}" phải viết hoa: "${word}".`, find: `${before}${mark}${space}${word}`, replacement: `${before}${mark}${space}${capitalize(word)}` });
  }
  return out;
}

function fixedNames(text: string): Finding[] {
  const out: Finding[] = [];
  const folded = fold(text);
  const taken: Array<[number, number]> = [];
  for (const name of FIXED_NAMES) {
    const pattern = new RegExp(`(?<![\\p{L}])${fold(name).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}(?![\\p{L}])`, "giu");
    for (const m of folded.matchAll(pattern)) {
      const start = m.index ?? 0;
      const end = start + m[0].length;
      if (taken.some(([a, b]) => start < b && end > a)) continue;
      taken.push([start, end]);
      const actual = text.slice(start, end);
      if (fold(actual) === fold(name)) continue;
      out.push({ ruleId: "cap.properName", message: `Viết hoa đúng tên "${name}" (đang viết "${actual}").`, find: actual, replacement: actual.startsWith("Uỷ") || actual.startsWith("uỷ") ? name.replace("Ủy", "Uỷ") : name });
    }
  }
  return out;
}

function documentKinds(text: string): Finding[] {
  const out: Finding[] = [];
  for (const m of text.matchAll(DOC_KIND)) {
    const word = m[0];
    if (!isLower(word[0])) continue;
    const next = text.slice((m.index ?? 0) + word.length).match(/^\s+\S+/)?.[0] ?? "";
    out.push({ ruleId: "cap.documentName", message: `Tên văn bản khi viện dẫn phải viết hoa chữ đầu: "${word}".`, find: `${word}${next}`, replacement: `${capitalize(word)}${next}` });
  }
  return out;
}

function references(text: string): Finding[] {
  const out: Finding[] = [];
  for (const m of text.matchAll(REFERENCE)) {
    const word = m[0];
    if (!isLower(word[0])) continue;
    const next = text.slice((m.index ?? 0) + word.length).match(/^\s+\S+/)?.[0] ?? "";
    out.push({ ruleId: "cap.reference", message: `Viện dẫn phần, chương, mục, điều, khoản, điểm phải viết hoa chữ đầu: "${word}${next}".`, find: `${word}${next}`, replacement: `${capitalize(word)}${next}` });
  }
  return out;
}

function checkable(p: CheckParagraph): boolean {
  const text = p.resolved.text;
  if (p.inTable || text.trim() === "") return false;
  const letters = text.replace(/[^\p{L}]/gu, "");
  return letters !== letters.toUpperCase();
}

export function checkCapitalization(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const reference = referenceOf(profile);
  const issues: Issue[] = [];
  for (const p of doc.paragraphs) {
    if (!checkable(p)) continue;
    const text = p.resolved.text;
    // Trích yếu, tên loại, tiêu đề (căn giữa) có quy cách riêng: trích yếu của Đảng viết thường, nên không xét đầu câu.
    const heading = p.resolved.align === "center" || /TrichYeu|TenLoai|TieuDe/i.test(p.styleName ?? "");
    const start = heading ? undefined : sentenceStart(text);
    const findings = [
      ...(start ? [start] : []),
      ...afterPeriod(text),
      ...fixedNames(text),
      ...documentKinds(text),
      ...(profile.headerKind === "quocHieu" ? references(text) : []),
    ];
    for (const f of findings) {
      issues.push(paragraphIssue(p, { ruleId: f.ruleId, severity: "warning", message: f.message, reference, actual: f.find.trim(), expected: f.replacement.trim(), fix: { kind: "edits", index: p.index, edits: [{ find: f.find, replacement: f.replacement }], label: `Viết hoa: ${f.replacement.trim()}` } }));
    }
  }
  return issues;
}
