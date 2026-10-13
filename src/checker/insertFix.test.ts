import { HD05 } from "../standards/hd05/profile";
import { checkPartyHeaderFooter } from "../standards/hd05/checks";
import { checkNd30HeaderFooter } from "../standards/nd30/checks";
import { ND30 } from "../standards/nd30/profile";
import { applyFix, fixIndex, isFormatOnly } from "./fixes";
import { planFixes, simulateFixAll } from "./fixAll";
import type { CheckDocument, CheckParagraph, Fix, Issue } from "./types";

interface P {
  text: string;
  body?: boolean;
  col?: number;
  row?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  sizePt?: number;
  align?: "center" | "left";
}

function doc(items: P[]): CheckDocument {
  const paragraphs: CheckParagraph[] = items.map((p, index) => ({
    index,
    inTable: !p.body,
    table: p.body ? undefined : { index: 0, row: p.row ?? 0, col: p.col ?? 0 },
    resolved: { kind: "paragraph", text: p.text, bold: p.bold, italic: p.italic, underline: p.underline, sizePt: p.sizePt ?? 14, align: p.align },
  }));
  return { pages: [], paragraphs };
}

const insertFix = (index: number, text: string): Fix => ({ kind: "insertAfter", index, text, align: "center", sizePt: 14, label: `thêm ${text}` });

describe("cách sửa thêm đoạn (insertAfter)", () => {
  it("chèn một đoạn ngay sau đoạn chỉ định, giữ ô bảng, dồn chỉ số các đoạn sau", () => {
    const d = doc([{ text: "A" }, { text: "B", col: 1 }, { text: "C", body: true }]);
    const next = applyFix(d, insertFix(1, "-----"));
    expect(next.paragraphs.map((p) => p.resolved.text)).toEqual(["A", "B", "-----", "C"]);
    expect(next.paragraphs.map((p) => p.index)).toEqual([0, 1, 2, 3]);
    expect(next.paragraphs[2]).toMatchObject({ inTable: true, table: { index: 0, row: 0, col: 1 }, resolved: { align: "center", sizePt: 14, bold: false } });
    expect(next.paragraphs[3].resolved.text).toBe("C");
  });

  it("không phải cách sửa chỉ định dạng; có chỉ số đoạn", () => {
    expect(isFormatOnly(insertFix(0, "x"))).toBe(false);
    expect(fixIndex(insertFix(4, "x"))).toBe(4);
  });

  it("kế hoạch đặt các cách sửa thêm đoạn sau cùng, chỉ số lớn trước, để chỉ số của các cách sửa khác không lệch", () => {
    const issue = (fix: Fix, ruleId: string): Issue => ({ ruleId, severity: "error", message: ruleId, reference: "r", location: { kind: "paragraph", index: 0, excerpt: "" }, fix });
    const plan = planFixes([
      issue(insertFix(1, "a"), "a"),
      issue({ kind: "font", index: 5, bold: true, label: "f" }, "f"),
      issue(insertFix(4, "b"), "b"),
      issue({ kind: "font", index: 0, bold: true, label: "g" }, "g"),
    ]);
    expect(plan.fixes.map((f) => (f.kind === "insertAfter" ? `i${f.index}` : `f${(f as { index: number }).index}`))).toEqual(["f0", "f5", "i4", "i1"]);
  });
});

describe("chuẩn Đảng: thêm đoạn còn thiếu", () => {
  const titleBlock = (over: P[] = []): P[] => [
    { text: "HƯỚNG DẪN", body: true, align: "center", bold: true, sizePt: 16 },
    { text: "thực hiện quy chế", body: true, align: "center", bold: true },
    ...over,
    { text: "Căn cứ Điều lệ Đảng;", body: true, align: "left" },
  ];
  const run = (d: CheckDocument): Issue[] => checkPartyHeaderFooter(d, HD05);

  it("thiếu dòng 5 dấu gạch nối dưới trích yếu: có cách sửa và sửa xong hết lỗi", () => {
    const d = doc(titleBlock());
    const issue = run(d).find((i) => i.ruleId === "title.dashes.missing");
    expect(issue?.fix).toMatchObject({ kind: "insertAfter", text: "-----", align: "center" });
    const fixed = applyFix(d, issue!.fix!);
    expect(run(fixed).map((i) => i.ruleId)).toEqual([]);
    expect(fixed.paragraphs.map((p) => p.resolved.text)).toEqual(["HƯỚNG DẪN", "thực hiện quy chế", "-----", "Căn cứ Điều lệ Đảng;"]);
  });

  it("thiếu dấu * dưới tên cơ quan ban hành: chèn sau dòng cuối của ô trái", () => {
    const d = doc([
      { text: "ĐẢNG CỘNG SẢN VIỆT NAM", col: 1, bold: true, sizePt: 15 },
      { text: "TỈNH ỦY ĐỒNG THÁP", col: 0 },
      { text: "BAN TỔ CHỨC", col: 0, bold: true },
    ]);
    const issue = run(d).find((i) => i.ruleId === "header.star");
    expect(issue?.fix).toMatchObject({ kind: "insertAfter", index: 2, text: "*" });
    const fixed = applyFix(d, issue!.fix!);
    expect(run(fixed)).toEqual([]);
    expect(fixed.paragraphs[3]).toMatchObject({ inTable: true, table: { col: 0 } });
  });
});

describe("khối cuối: các lỗi trước đây chỉ báo nay có cách sửa", () => {
  const footer = (over: Record<string, Partial<P>> = {}): P[] =>
    [
      { key: "nn", text: "Nơi nhận:", bold: true, italic: true, sizePt: 12 },
      { key: "a", text: "- Như trên;", sizePt: 11 },
      { key: "b", text: "- Lưu: VT.", sizePt: 11 },
      { key: "qh", text: "KT. CHI CỤC TRƯỞNG", col: 1, bold: true },
      { key: "ten", text: "Mai Hiếu Hiền", col: 1, bold: true },
    ].map(({ key, ...p }) => ({ row: 0, ...p, ...over[key] }));
  const run = (d: CheckDocument): Issue[] => checkNd30HeaderFooter(d, ND30);

  it("mục nơi nhận thiếu gạch đầu dòng được thêm", () => {
    const d = doc(footer({ a: { text: "Như trên;" } }));
    expect(run(d).map((i) => i.ruleId)).toEqual(["footer.noiNhan.dash"]);
    const result = simulateFixAll(d, run);
    expect(result.remaining).toEqual([]);
    expect(result.doc.paragraphs[1].resolved.text).toBe("- Như trên;");
  });

  it("họ tên in hoa toàn bộ được đổi sang viết hoa chữ cái đầu mỗi chữ", () => {
    const result = simulateFixAll(doc(footer({ ten: { text: "NGUYỄN VĂN ĐỨC" } })), run);
    expect(result.remaining).toEqual([]);
    expect(result.doc.paragraphs[4].resolved.text).toBe("Nguyễn Văn Đức");
  });

  it("học hàm, học vị ở đầu họ tên được bỏ", () => {
    for (const [raw, expected] of [["TS. Nguyễn Văn A", "Nguyễn Văn A"], ["PGS.TS Lê Thị B", "Lê Thị B"], ["Thạc sĩ Trần C", "Trần C"], ["BS CKII Phạm D", "Phạm D"]] as const) {
      const result = simulateFixAll(doc(footer({ ten: { text: raw } })), run);
      expect(result.remaining.filter((i) => i.fix)).toEqual([]);
      expect(result.doc.paragraphs[4].resolved.text).toBe(expected);
    }
  });
});
