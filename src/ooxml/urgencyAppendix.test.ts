import { checkUrgencyMarks } from "../checker/urgencyRules";
import { applyFix } from "../checker/fixes";
import type { CheckDocument, Issue } from "../checker/types";
import { extractDocModel, type ParagraphModel, type TableModel } from "../rules/docModel";
import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { appendixNumberingXml, buildAppendix, referenceText, validateAppendix } from "./appendix";
import { expectMatchesGolden } from "./__golden__/golden";
import { buildUrgencyMark, markSize, validateUrgencyMark } from "./urgencyMark";

function blocks(xml: string) {
  return extractDocModel({ documentXml: `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${xml}</w:body></w:document>` }).blocks;
}

describe("khung dấu chỉ mức độ khẩn", () => {
  it.each(ND30.urgencyMarks.map((mark) => [mark.text, mark] as const))("%s: bảng một ô, viền đơn, in hoa đậm, đúng kích thước", (_name, mark) => {
    const xml = buildUrgencyMark({ text: mark.text, widthMm: mark.widthMm, heightMm: mark.heightMm, sizePt: 14 });
    const [table, tail] = blocks(xml) as [TableModel, ParagraphModel];
    expect(table.kind).toBe("table");
    expect(table.rows).toHaveLength(1);
    expect(table.rows[0]).toHaveLength(1);
    expect(table.rows[0][0][0]).toMatchObject({ text: mark.text, bold: true, align: "center", sizePt: 14 });
    expect(table.bordersNone).toBe(false);
    expect(table.columnsMm[0]).toBeCloseTo(mark.widthMm!, 0);
    expect(xml).toContain('w:hRule="exact"');
    expect(xml).toContain(`<w:trHeight w:val="${Math.round((mark.heightMm! * 1440) / 25.4)}"`);
    expect(xml).toContain('<w:vAlign w:val="center"/>');
    expect(tail.kind).toBe("paragraph");
  });

  it("chữ nhập thường được in hoa; ký tự đặc biệt được thoát", () => {
    expect(buildUrgencyMark({ text: " hỏa  tốc ", sizePt: 14 })).toContain("HỎA TỐC");
    expect(buildUrgencyMark({ text: "A&B", sizePt: 14 })).toContain("A&amp;B");
  });

  it("không nêu kích thước thì rộng vừa chữ, cao 8 mm", () => {
    const size = markSize({ text: "LƯU HÀNH NỘI BỘ", sizePt: 12 });
    expect(size.heightMm).toBe(8);
    expect(size.widthMm).toBeGreaterThan(30);
  });

  it("đầu vào sai bị từ chối", () => {
    expect(validateUrgencyMark({ text: " ", sizePt: 14 })).not.toEqual([]);
    expect(validateUrgencyMark({ text: "KHẨN", sizePt: 14, widthMm: 0 })).not.toEqual([]);
    expect(() => buildUrgencyMark({ text: "", sizePt: 14 })).toThrow();
  });

  it("căn khung trong vùng chữ", () => {
    expect(buildUrgencyMark({ text: "KHẨN", sizePt: 14, align: "right" })).toMatch(/<w:tblPr><w:tblW w:w="\d+" w:type="dxa"\/><w:jc w:val="right"\/>/);
    expect(buildUrgencyMark({ text: "KHẨN", sizePt: 14 })).not.toContain('<w:jc w:val="right"/>');
  });

  it("golden", () => {
    expectMatchesGolden("urgencyMark-khan.xml", buildUrgencyMark({ text: "KHẨN", widthMm: 20, heightMm: 8, sizePt: 14 }));
  });
});

describe("trang phụ lục", () => {
  const base = { title: "Danh sách đại biểu", reference: "Kèm theo Công văn số 05/CV-VP ngày 05 tháng 01 năm 2026", number: 1, stylePrefix: "ND30_" };

  it("nhãn đánh số tự động, tên in hoa, chỉ dẫn trong ngoặc đơn", () => {
    const xml = buildAppendix(base);
    const paragraphs = blocks(xml) as ParagraphModel[];
    expect(paragraphs.map((p) => p.styleId)).toEqual(["ND30PhuLucSo", "ND30PhuLucTieuDe", "ND30PhuLucChiDan"]);
    expect(paragraphs[1].text).toBe("DANH SÁCH ĐẠI BIỂU");
    expect(paragraphs[2].text).toBe("(Kèm theo Công văn số 05/CV-VP ngày 05 tháng 01 năm 2026)");
  });

  it("nhãn bắt đầu trang mới", () => {
    expect(buildAppendix(base)).toContain('<w:pStyle w:val="ND30PhuLucSo"/><w:pageBreakBefore/><w:numPr>');
  });

  it("số thứ tự La Mã nối tiếp theo phụ lục đã có", () => {
    const xml = appendixNumberingXml(3);
    expect(xml).toContain('<w:start w:val="3"/>');
    expect(xml).toContain('<w:numFmt w:val="upperRoman"/>');
    expect(xml).toContain('<w:lvlText w:val="Phụ lục %1"/>');
    expect(xml).toContain('<w:suff w:val="nothing"/>');
  });

  it("không có chỉ dẫn thì bỏ đoạn đó; có chỗ nội dung thì thêm đoạn nội dung", () => {
    expect(blocks(buildAppendix({ ...base, reference: "" }))).toHaveLength(2);
    expect(blocks(buildAppendix({ ...base, withBody: true })).map((b) => (b as ParagraphModel).styleId).pop()).toBe("ND30NoiDung");
  });

  it("đầu vào sai bị từ chối", () => {
    expect(validateAppendix({ ...base, title: " " })).not.toEqual([]);
    expect(validateAppendix({ ...base, number: 0 })).not.toEqual([]);
    expect(() => buildAppendix({ ...base, number: 100 })).toThrow();
    expect(referenceText("(Đã có)")).toBe("(Đã có)");
    expect(referenceText(" ")).toBe("");
  });

  it("golden", () => {
    expectMatchesGolden("appendix-1.xml", buildAppendix(base));
  });
});

describe("nhận biết dấu khẩn trong kiểm tra", () => {
  function doc(text: string, over: Partial<ParagraphModel> = {}, inTable = true): CheckDocument {
    return { pages: [], paragraphs: [{ index: 0, inTable, table: inTable ? { index: 0, row: 0, col: 0 } : undefined, resolved: { kind: "paragraph", text, bold: true, sizePt: 14, ...over } }] };
  }
  const ids = (issues: Issue[]) => issues.map((i) => i.ruleId);

  it("dấu đúng thì không lỗi (cả hai chuẩn)", () => {
    expect(checkUrgencyMarks(doc("HỎA TỐC"), ND30)).toEqual([]);
    expect(checkUrgencyMarks(doc("THƯỢNG KHẨN"), HD05)).toEqual([]);
  });

  it("viết thường, không đậm, sai cỡ bị báo và sửa được", () => {
    const d = doc("Khẩn", { bold: false, sizePt: 18 });
    const issues = checkUrgencyMarks(d, ND30);
    expect(ids(issues)).toEqual(["urgency.upper", "urgency.bold", "urgency.size"]);
    let fixed = d;
    for (const issue of issues) fixed = applyFix(fixed, issue.fix!);
    expect(checkUrgencyMarks(fixed, ND30)).toEqual([]);
  });

  it("đoạn khác hoặc đoạn ngoài bảng không bị xét", () => {
    expect(checkUrgencyMarks(doc("Rất khẩn trương thực hiện"), ND30)).toEqual([]);
    expect(checkUrgencyMarks(doc("khẩn", {}, false), ND30)).toEqual([]);
  });
});
