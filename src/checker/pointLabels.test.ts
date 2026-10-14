import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { diemLabel, VIETNAMESE_LETTERS } from "../ooxml/numbering";
import { simulateFixAll } from "./fixAll";
import { checkPointLabels } from "./pointLabels";
import type { CheckDocument, CheckParagraph } from "./types";

function doc(texts: string[]): CheckDocument {
  const paragraphs: CheckParagraph[] = texts.map((text, index) => ({ index, inTable: false, resolved: { kind: "paragraph", text, align: "both", sizePt: 14 } }));
  return { pages: [], paragraphs };
}

describe("chữ cái tiếng Việt của Điểm", () => {
  it("thứ tự a, b, c, d, đ, e, g, h, i, k…; không có f, j, w, z", () => {
    expect(VIETNAMESE_LETTERS.slice(0, 10)).toEqual(["a", "b", "c", "d", "đ", "e", "g", "h", "i", "k"]);
    expect(VIETNAMESE_LETTERS).not.toEqual(expect.arrayContaining(["f", "j", "w", "z"]));
    expect(diemLabel(4)).toBe("đ)");
    expect(diemLabel(5)).toBe("e)");
  });
});

describe("kiểm tra nhãn Điểm", () => {
  it("dãy đúng chữ cái tiếng Việt không bị báo", () => {
    expect(checkPointLabels(doc(["a) Một;", "b) Hai;", "c) Ba;", "d) Bốn;", "đ) Năm;", "e) Sáu."]), ND30)).toEqual([]);
  });

  it("dãy theo bảng chữ Latin (thiếu đ) bị báo và sửa nhãn về đúng thứ tự", () => {
    const d = doc(["a) Một;", "b) Hai;", "c) Ba;", "d) Bốn;", "e) Năm;", "f) Sáu;", "g) Bảy."]);
    const issues = checkPointLabels(d, ND30);
    expect(issues.map((i) => i.actual)).toEqual(["e)", "f)"]);
    expect(issues.every((i) => i.severity === "error" && i.fix?.kind === "edits")).toBe(true);
    const fixed = simulateFixAll(d, (x) => checkPointLabels(x, ND30));
    expect(fixed.remaining).toEqual([]);
    expect(fixed.doc.paragraphs.map((p) => p.resolved.text.slice(0, 2))).toEqual(["a)", "b)", "c)", "d)", "đ)", "e)", "g)"]);
  });

  it("dãy mới bắt đầu lại sau một đoạn không có nhãn (Khoản khác)", () => {
    expect(checkPointLabels(doc(["a) Một;", "b) Hai.", "2. Khoản hai:", "a) Một;", "b) Hai."]), HD05)).toEqual([]);
  });

  it("không xét đoạn trong bảng và đoạn không bắt đầu bằng a)", () => {
    expect(checkPointLabels(doc(["c) Lẻ;", "e) Lẻ."]), ND30)).toEqual([]);
  });
});
