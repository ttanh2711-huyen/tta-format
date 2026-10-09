import { ND30 } from "../standards/nd30/profile";
import { HD05 } from "../standards/hd05/profile";
import type { StandardProfile } from "../standards/types";
import { checkNd30HeaderFooter } from "../standards/nd30/checks";
import { runChecks } from "./engine";
import { formatReport } from "./exportReport";
import { planFixes, simulateFixAll } from "./fixAll";
import type { CheckDocument, CheckParagraph, Fix, Issue, PageInfo } from "./types";

const where = { kind: "paragraph", index: 0, excerpt: "x" } as const;
const issue = (ruleId: string, fix?: Fix, over: Partial<Issue> = {}): Issue => ({ ruleId, severity: "error", message: ruleId, reference: "ref", location: where, fix, ...over });
const font = (index: number, patch: Partial<Extract<Fix, { kind: "font" }>>): Fix => ({ kind: "font", index, label: `font ${index}`, ...patch });

describe("planFixes: gom và sắp xếp", () => {
  it("gộp các cách sửa font cùng một đoạn thành một", () => {
    const plan = planFixes([issue("a", font(3, { bold: true })), issue("b", font(3, { sizePt: 14 }))]);
    expect(plan.fixes).toEqual([{ kind: "font", index: 3, bold: true, sizePt: 14, label: "font 3" }]);
    expect(plan.changes[0].ruleIds).toEqual(["a", "b"]);
    expect(plan.manual).toEqual([]);
  });

  it("gộp lề của cùng một phần, giữ lề của phần khác riêng", () => {
    const page = (section: number, side: "top" | "left", mm: number): Fix => ({ kind: "page", section, marginsMm: { [side]: mm }, label: `${side}` });
    const plan = planFixes([issue("a", page(1, "top", 20)), issue("b", page(1, "left", 30)), issue("c", page(2, "top", 20))]);
    expect(plan.fixes).toHaveLength(2);
    expect(plan.fixes[0]).toMatchObject({ kind: "page", section: 1, marginsMm: { top: 20, left: 30 } });
  });

  it("bỏ cách sửa trùng hệt nhau", () => {
    const f = font(1, { bold: true });
    expect(planFixes([issue("a", f), issue("a", { ...f })]).fixes).toHaveLength(1);
  });

  it("hai cách thay chữ khác nhau trên một đoạn: giữ cái đầu, hoãn cái sau", () => {
    const r1: Fix = { kind: "replace", index: 2, find: "ab", replacement: "AB", label: "1" };
    const r2: Fix = { kind: "replace", index: 2, find: "ab", replacement: "ab.", label: "2" };
    const second = issue("b", r2);
    const plan = planFixes([issue("a", r1), second]);
    expect(plan.fixes).toEqual([r1]);
    expect(plan.manual).toEqual([second]);
  });

  it("lỗi không có cách sửa nằm trong manual", () => {
    const none = issue("x");
    expect(planFixes([none]).manual).toEqual([none]);
  });

  it("includeText=false loại cách sửa thay chữ; skipRule loại theo quy tắc", () => {
    const rep = issue("t", { kind: "replace", index: 0, find: "a", replacement: "b", label: "r" });
    const fnt = issue("font.size", font(1, { sizePt: 14 }));
    expect(planFixes([rep, fnt], { includeText: false }).fixes).toEqual([fnt.fix]);
    expect(planFixes([rep, fnt], { skipRule: (id) => id.startsWith("font.") }).fixes).toEqual([rep.fix]);
  });

  it("sắp xếp: trang trước, rồi theo chỉ số đoạn, cùng đoạn thì font trước thay chữ trước Exactly", () => {
    const page: Fix = { kind: "page", section: 1, paper: "A4", label: "p" };
    const exact: Fix = { kind: "exact", index: 1, pt: 20, label: "e" };
    const rep: Fix = { kind: "replace", index: 1, find: "a", replacement: "b", label: "r" };
    const plan = planFixes([issue("e", exact), issue("f", font(5, { bold: true })), issue("r", rep), issue("p", page), issue("g", font(1, { bold: true }))]);
    expect(plan.fixes.map((f) => f.kind)).toEqual(["page", "font", "replace", "exact", "font"]);
  });
});

describe("simulateFixAll", () => {
  const page = (profile: StandardProfile, over: Partial<PageInfo> = {}): PageInfo => ({
    widthMm: 210,
    heightMm: 297,
    marginsMm: { ...profile.pageDefaults.marginsMm },
    headerDistanceMm: profile.pageDefaults.headerDistanceMm,
    footerDistanceMm: profile.pageDefaults.footerDistanceMm,
    mirrorMargins: false,
    ...over,
  });
  const LONG = "Thực hiện Thông báo số 458/TB-VPUB của Văn phòng Ủy ban nhân dân thành phố về tập huấn bồi dưỡng.";
  const bad = (index: number): CheckParagraph => ({
    index,
    inTable: false,
    resolved: { kind: "paragraph", text: LONG, align: "left", firstLineIndentMm: 0, spaceBeforePt: 0, spaceAfterPt: 0, lineMultiple: 2, fontName: "Arial", sizePt: 9 },
  });

  it.each([
    ["Chính quyền", ND30],
    ["Đảng", HD05],
  ])("sửa hết lỗi sửa được (%s) chỉ trong một lượt khi không có chồng chéo", (_name, profile) => {
    const doc: CheckDocument = { pages: [page(profile, { marginsMm: { top: 5, bottom: 5, left: 5, right: 5 } })], paragraphs: [bad(0), bad(1)] };
    const result = simulateFixAll(doc, (d) => runChecks(d, profile));
    expect(result.rounds).toBe(1);
    expect(result.remaining.filter((i) => i.fix)).toEqual([]);
    expect(result.remaining).toEqual([]);
  });

  it("lỗi chồng chéo trên một đoạn (hai lần thay chữ) xong sau lượt thứ hai", () => {
    // Chức vụ viết thường và quyền hạn "TM" thiếu dấu chấm cùng cần thay cả dòng chữ.
    const cell = (index: number, col: number, text: string, bold = false): CheckParagraph => ({
      index,
      inTable: true,
      table: { index: 0, row: 0, col },
      resolved: { kind: "paragraph", text, bold, sizePt: 13 },
    });
    const doc: CheckDocument = {
      pages: [],
      paragraphs: [cell(0, 0, "Nơi nhận:"), cell(1, 1, "TM Chủ tịch", true), cell(2, 1, "Nguyễn Văn A", true)],
    };
    const result = simulateFixAll(doc, (d) => checkNd30HeaderFooter(d, ND30));
    expect(result.rounds).toBeGreaterThanOrEqual(2);
    expect(result.remaining.filter((i) => i.fix)).toEqual([]);
  });
});

describe("formatReport", () => {
  const header = { standardName: "Chính quyền", checkedAt: new Date(2026, 9, 9, 14, 0, 0) };

  it("không lỗi thì ghi rõ", () => {
    expect(formatReport([], header)).toContain("Không thấy lỗi.");
  });

  it("liệt kê từng lỗi kèm vị trí, căn cứ, cách sửa", () => {
    const text = formatReport(
      [issue("a", font(0, { bold: true }), { message: "Thiếu đậm", actual: "không", expected: "đậm" }), issue("b", undefined, { severity: "warning", message: "Xem lại" })],
      header,
    );
    expect(text).toContain("CHUẨN CHÍNH QUYỀN");
    expect(text).toContain("1 lỗi, 1 cảnh báo; 1 mục sửa được tự động.");
    expect(text).toContain("1. [Lỗi] Thiếu đậm");
    expect(text).toContain("Hiện có: không | Cần: đậm");
    expect(text).toContain("Cách sửa: font 0");
    expect(text).toContain("2. [Cảnh báo] Xem lại");
    expect(text).toContain("Cách sửa: cần sửa thủ công");
  });
});
