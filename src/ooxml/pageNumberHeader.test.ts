import { hasPageField, readJustification } from "./facts";
import { pageNumberParagraph, planPageNumbers, readPageNumberState, type SectionState } from "./pageNumberHeader";

const opts = { sizePt: 14, hideFirst: true };

describe("pageNumberParagraph", () => {
  it("đoạn căn giữa, Times New Roman, cỡ theo tuỳ chọn, chứa trường PAGE", () => {
    const xml = pageNumberParagraph(14);
    expect(hasPageField(xml)).toBe(true);
    expect(readJustification(xml)).toBe("center");
    expect(xml).toContain('<w:sz w:val="28"/>');
    expect(xml).toContain("Times New Roman");
    expect(pageNumberParagraph(13)).toContain('<w:sz w:val="26"/>');
  });

  it("không có khoảng giãn thừa làm đẩy nội dung trang", () => {
    expect(pageNumberParagraph(14)).toContain('w:before="0" w:after="0"');
  });
});

describe("readPageNumberState: đọc header Word trả về", () => {
  it("nhận ra header đã có số trang đúng cỡ và căn giữa", () => {
    expect(readPageNumberState(`<w:hdr>${pageNumberParagraph(14)}</w:hdr>`)).toEqual({ hasPage: true, centered: true, sizePt: 14 });
  });

  it("nhận ra trường PAGE dạng phức hợp do Word tự tạo", () => {
    const xml = '<w:hdr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE   \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>';
    expect(readPageNumberState(xml)).toEqual({ hasPage: true, centered: true, sizePt: 14 });
  });

  it("header trống hoặc không có số trang", () => {
    expect(readPageNumberState("<w:hdr><w:p/></w:hdr>")).toEqual({ hasPage: false, centered: false, sizePt: undefined });
    expect(readPageNumberState('<w:hdr><w:p><w:r><w:t>Tên cơ quan</w:t></w:r></w:p></w:hdr>').hasPage).toBe(false);
  });

  it("NUMPAGES không được tính là số trang", () => {
    expect(readPageNumberState('<w:hdr><w:p><w:fldSimple w:instr=" NUMPAGES "/></w:p></w:hdr>').hasPage).toBe(false);
  });
});

const state = (over: Partial<SectionState> = {}): SectionState => ({
  differentFirst: true,
  primary: { hasPage: true, centered: true, sizePt: 14 },
  first: { hasPage: false, empty: true },
  ...over,
});

describe("planPageNumbers", () => {
  it("tài liệu đã đúng thì không có việc gì (idempotent)", () => {
    expect(planPageNumbers([state()], opts)).toEqual([{ setDifferentFirst: undefined, writePrimary: false, clearFirst: false }]);
  });

  it("chưa có số trang: ghi header chính và bật trang đầu khác biệt", () => {
    const plan = planPageNumbers([state({ differentFirst: false, primary: { hasPage: false, centered: false, sizePt: undefined } })], opts);
    expect(plan).toEqual([{ setDifferentFirst: true, writePrimary: true, clearFirst: false }]);
  });

  it("số trang sai cỡ hoặc không căn giữa thì ghi lại", () => {
    expect(planPageNumbers([state({ primary: { hasPage: true, centered: true, sizePt: 12 } })], opts)[0].writePrimary).toBe(true);
    expect(planPageNumbers([state({ primary: { hasPage: true, centered: false, sizePt: 14 } })], opts)[0].writePrimary).toBe(true);
  });

  it("trang 1 phải trống: xoá số trang nếu header trang đầu đang có trường PAGE, giữ nguyên nội dung khác", () => {
    expect(planPageNumbers([state({ first: { hasPage: true, empty: false } })], opts)[0].clearFirst).toBe(true);
    expect(planPageNumbers([state({ first: { hasPage: false, empty: false } })], opts)[0].clearFirst).toBe(false);
  });

  it("hiện số ở trang 1: tắt trang đầu khác biệt", () => {
    const plan = planPageNumbers([state({ differentFirst: true })], { sizePt: 14, hideFirst: false });
    expect(plan[0]).toEqual({ setDifferentFirst: false, writePrimary: false, clearFirst: false });
  });

  it("nhiều phần: chỉ phần đầu ẩn số trang 1; phần sau không bật trang đầu khác biệt", () => {
    const plan = planPageNumbers(
      [state(), state({ differentFirst: true }), state({ differentFirst: false, primary: { hasPage: false, centered: false, sizePt: undefined } })],
      opts,
    );
    expect(plan[0].setDifferentFirst).toBeUndefined();
    expect(plan[1].setDifferentFirst).toBe(false);
    expect(plan[2]).toEqual({ setDifferentFirst: undefined, writePrimary: true, clearFirst: false });
  });

  it("phần sau không xoá header trang đầu của mình", () => {
    expect(planPageNumbers([state(), state({ first: { hasPage: true, empty: false } })], opts)[1].clearFirst).toBe(false);
  });
});
