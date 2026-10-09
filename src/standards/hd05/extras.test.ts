import { applyFix } from "../../checker/fixes";
import type { CheckDocument, Issue } from "../../checker/types";
import { checkUrgencyMarks } from "../../checker/urgencyRules";
import { expectMatchesGolden } from "../../ooxml/__golden__/golden";
import { extractDocModel, type ParagraphModel, type TableModel } from "../../rules/docModel";
import { buildCirculationMark, buildContactInfo, buildDraftNote, buildNoPublicNote, buildPartyAppendix, checkPartyExtras, NO_PUBLIC_NOTE, partyAppendixNumberingXml, validatePartyAppendix } from "./extras";
import { HD05 } from "./profile";

const blocks = (xml: string) => extractDocModel({ documentXml: `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${xml}</w:body></w:document>` }).blocks;

describe("chỉ dẫn phạm vi lưu hành", () => {
  it("khung viền đơn, canh phải, 12 pt, in hoa đậm", () => {
    const [table] = blocks(buildCirculationMark("xem xong trả lại", 12)) as [TableModel];
    expect(table.bordersNone).toBe(false);
    expect(table.rows[0][0][0]).toMatchObject({ text: "XEM XONG TRẢ LẠI", bold: true, sizePt: 12 });
    expect(buildCirculationMark("xem xong trả lại", 12)).toContain('<w:jc w:val="right"/>');
  });

  it("được nhận biết như dấu khẩn khi trùng một dấu của chuẩn (KHẨN) và kiểm tra cỡ chữ", () => {
    const doc: CheckDocument = { pages: [], paragraphs: [{ index: 0, inTable: true, table: { index: 0, row: 0, col: 0 }, resolved: { kind: "paragraph", text: "KHẨN", bold: true, sizePt: 12 } }] };
    expect(checkUrgencyMarks(doc, HD05).map((i) => i.ruleId)).toEqual(["urgency.size"]);
  });
});

describe("chỉ dẫn dự thảo", () => {
  it.each(["Dự thảo", "Dự thảo lần 2", "dự thảo lần 12"])("%s hợp lệ, nghiêng, cỡ 14", (text) => {
    const [p] = blocks(buildDraftNote(text, 14)) as [ParagraphModel];
    expect(p).toMatchObject({ text, italic: true, sizePt: 14, align: "left" });
  });

  it("dạng khác bị từ chối", () => {
    expect(() => buildDraftNote("Bản nháp", 14)).toThrow();
    expect(() => buildDraftNote("Dự thảo lần hai", 14)).toThrow();
  });
});

describe("dòng không phổ biến và thông tin liên hệ", () => {
  it("dòng không phổ biến cỡ 12, canh giữa", () => {
    const [p] = blocks(buildNoPublicNote(12)) as [ParagraphModel];
    expect(p).toMatchObject({ text: NO_PUBLIC_NOTE, sizePt: 12, align: "center" });
  });

  it("thông tin liên hệ: đường kẻ ngang rồi các dòng 10 pt, bỏ dòng trống", () => {
    const xml = buildContactInfo(["Số 1 Nguyễn Huệ", "", "ĐT: 0123"], 10);
    const paragraphs = blocks(xml) as ParagraphModel[];
    expect(paragraphs.map((p) => p.text)).toEqual(["", "Số 1 Nguyễn Huệ", "ĐT: 0123"]);
    expect(paragraphs.slice(1).every((p) => p.sizePt === 10)).toBe(true);
    expect(xml).toContain("<w:pBdr><w:top");
    expect(() => buildContactInfo([" "], 10)).toThrow();
  });

  it("golden", () => {
    expectMatchesGolden("partyContact.xml", buildContactInfo(["Số 1 Nguyễn Huệ", "ĐT: 0123"], 10));
  });
});

describe("trang phụ lục của Đảng", () => {
  const base = { title: "Danh sách đại biểu", reference: "Kèm theo Hướng dẫn số 05-HD/VPTW", number: 2, sizePt: 14 };

  it("bắt đầu trang mới, nhãn tự đánh số bắt đầu từ số của phụ lục, tên in hoa đậm, chỉ dẫn nghiêng", () => {
    const xml = buildPartyAppendix(base);
    expect(xml).toContain("<w:pageBreakBefore/><w:numPr>");
    expect(partyAppendixNumberingXml(2)).toContain('<w:start w:val="2"/>');
    expect(partyAppendixNumberingXml(2)).toContain('<w:lvlText w:val="Phụ lục %1"/>');
    const paragraphs = blocks(xml) as ParagraphModel[];
    expect(paragraphs[1]).toMatchObject({ text: "DANH SÁCH ĐẠI BIỂU", bold: true, align: "center" });
    expect(paragraphs[2]).toMatchObject({ text: "(Kèm theo Hướng dẫn số 05-HD/VPTW)", italic: true });
  });

  it("đầu vào sai bị từ chối", () => {
    expect(validatePartyAppendix({ ...base, title: " " })).not.toEqual([]);
    expect(() => buildPartyAppendix({ ...base, number: 0 })).toThrow();
  });

  it("golden", () => {
    expectMatchesGolden("partyAppendix-2.xml", buildPartyAppendix(base));
  });
});

describe("kiểm tra nhận biết", () => {
  const doc = (text: string, over: Partial<ParagraphModel> = {}): CheckDocument => ({ pages: [], paragraphs: [{ index: 0, inTable: false, resolved: { kind: "paragraph", text, italic: true, sizePt: 14, ...over } }] });
  const ids = (issues: Issue[]) => issues.map((i) => i.ruleId);

  it("chỉ dẫn dự thảo đúng thì không lỗi; thiếu nghiêng hoặc sai cỡ thì báo và sửa được", () => {
    expect(checkPartyExtras(doc("Dự thảo lần 2"), HD05)).toEqual([]);
    const bad = doc("Dự thảo lần 2", { italic: false, sizePt: 12 });
    const issues = checkPartyExtras(bad, HD05);
    expect(ids(issues)).toEqual(["extras.draft.italic", "extras.draft.size"]);
    let fixed = bad;
    for (const issue of issues) fixed = applyFix(fixed, issue.fix!);
    expect(checkPartyExtras(fixed, HD05)).toEqual([]);
  });

  it("dòng không phổ biến sai cỡ bị báo", () => {
    expect(ids(checkPartyExtras(doc(NO_PUBLIC_NOTE, { sizePt: 14 }), HD05))).toEqual(["extras.noPublic.size"]);
    expect(checkPartyExtras(doc(NO_PUBLIC_NOTE, { sizePt: 12 }), HD05)).toEqual([]);
  });

  it("đoạn khác không bị xét", () => {
    expect(checkPartyExtras(doc("Dự thảo nghị quyết về đất đai"), HD05)).toEqual([]);
  });
});
