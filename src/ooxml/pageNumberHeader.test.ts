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
  const pkg = (body: string) =>
    `<pkg:package><pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:styleId="X"><w:pPr><w:jc w:val="left"/></w:pPr></w:style></w:styles></pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${body}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;

  it("chỉ đọc phần thân của gói, không đọc nhầm cỡ chữ và căn lề của phần styles đứng trước", () => {
    expect(readPageNumberState(pkg(pageNumberParagraph(14)))).toEqual({ hasPage: true, centered: true, sizePt: 14 });
  });

  it("Word bỏ cỡ chữ trực tiếp khi bằng cỡ mặc định của tài liệu: lấy cỡ docDefaults", () => {
    const noSize = `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>`;
    const withDefault = (half: number) => pkg(noSize).replace('<w:sz w:val="22"/>', `<w:sz w:val="${half}"/>`);
    expect(readPageNumberState(withDefault(28))).toEqual({ hasPage: true, centered: true, sizePt: 14 });
    expect(readPageNumberState(`<w:hdr>${noSize}</w:hdr>`).sizePt).toBeUndefined();
  });

  it("nhận ra header đã có số trang đúng cỡ và căn giữa", () => {
    expect(readPageNumberState(`<w:hdr>${pageNumberParagraph(14)}</w:hdr>`)).toEqual({ hasPage: true, centered: true, sizePt: 14 });
  });

  it("đánh giá đúng đoạn chứa PAGE dù có đoạn trống khác cỡ, khác căn lề đứng trước hoặc sau", () => {
    const empty = '<w:p><w:pPr><w:jc w:val="left"/><w:rPr><w:sz w:val="22"/></w:rPr></w:pPr></w:p>';
    expect(readPageNumberState(`<w:hdr>${empty}${pageNumberParagraph(14)}${empty}</w:hdr>`)).toEqual({ hasPage: true, centered: true, sizePt: 14 });
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
