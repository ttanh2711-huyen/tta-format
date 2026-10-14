import { diemLabel, VIETNAMESE_LETTERS } from "../ooxml/numbering";
import type { StandardProfile } from "../standards/types";
import { paragraphIssue } from "./common";
import type { CheckDocument, Issue } from "./types";

const LABEL = /^(\s*)([\p{Ll}])\)(?=\s)/u;

/**
 * Nhãn Điểm gõ tay phải theo thứ tự chữ cái tiếng Việt (a, b, c, d, đ, e, g…). Xét từng dãy đoạn liền nhau ngoài bảng
 * bắt đầu bằng "a)"; đoạn không có nhãn kết thúc dãy (sang Khoản khác). Danh sách đánh số tự động của Word không xét
 * được ở đây vì nhãn không nằm trong chữ của đoạn.
 */
export function checkPointLabels(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const reference = profile.headerKind === "quocHieu" ? "Nghị định 30/2020/NĐ-CP, Phụ lục I: điểm ký hiệu bằng chữ cái tiếng Việt theo thứ tự (a, b, c, d, đ, e…)" : "Hướng dẫn 05: điểm ký hiệu bằng chữ cái tiếng Việt theo thứ tự (a, b, c, d, đ, e…)";
  const issues: Issue[] = [];
  let position = -1;
  for (const p of doc.paragraphs) {
    const m = p.inTable ? null : LABEL.exec(p.resolved.text);
    if (!m) {
      if (p.resolved.text.trim() !== "") position = -1;
      continue;
    }
    const letter = m[2];
    if (position < 0) {
      if (letter !== "a") continue;
      position = 0;
      continue;
    }
    position += 1;
    const expected = VIETNAMESE_LETTERS[position];
    if (expected === undefined || letter === expected) continue;
    const actual = `${letter})`;
    issues.push(
      paragraphIssue(p, {
        ruleId: "list.diem.order",
        severity: "error",
        message: `Điểm thứ ${position + 1} phải là "${diemLabel(position)}", đang là "${actual}" (thứ tự chữ cái tiếng Việt có "đ", không có f, j, w, z).`,
        reference,
        actual,
        expected: diemLabel(position),
        fix: { kind: "edits", index: p.index, edits: [{ find: `${m[1]}${actual}`, replacement: `${m[1]}${diemLabel(position)}` }], label: `Đổi nhãn ${actual} thành ${diemLabel(position)}` },
      }),
    );
  }
  return issues;
}
