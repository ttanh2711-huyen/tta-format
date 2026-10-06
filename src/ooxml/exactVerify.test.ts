import { countExactParagraphs, verifyExactPass } from "./exactVerify";

const doc = (paragraphs: string[]) =>
  `<pkg:package><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${paragraphs.join("")}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
const exact = (text: string) => `<w:p><w:pPr><w:spacing w:line="400" w:lineRule="exact"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const auto = (text: string) => `<w:p><w:pPr><w:spacing w:line="280" w:lineRule="auto"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

describe("countExactParagraphs", () => {
  it("đếm đoạn có giãn dòng Exactly đúng giá trị twips", () => {
    expect(countExactParagraphs(doc([exact("a"), auto("b"), exact("c")]), 400)).toBe(2);
  });

  it("không tính Exactly sai giá trị và không tính style khác phần thân", () => {
    const other = `<w:p><w:pPr><w:spacing w:line="440" w:lineRule="exact"/></w:pPr></w:p>`;
    expect(countExactParagraphs(doc([other]), 400)).toBe(0);
  });
});

describe("verifyExactPass", () => {
  const before = ["một", "hai", "ba"];

  it("đạt khi số đoạn, chữ và số đoạn Exactly đều khớp", () => {
    const result = verifyExactPass({ before, after: before, ooxml: doc([exact("một"), exact("hai"), exact("ba")]), twips: 400, expectedExact: 3 });
    expect(result).toEqual({ ok: true });
  });

  it("không đạt khi số đoạn thay đổi (chèn thêm đoạn trống)", () => {
    const result = verifyExactPass({ before, after: [...before, ""], ooxml: doc([exact("một")]), twips: 400, expectedExact: 1 });
    expect(result).toEqual({ ok: false, problem: "số đoạn đổi từ 3 thành 4" });
  });

  it("không đạt khi chữ bị đổi", () => {
    const result = verifyExactPass({ before, after: ["một", "HAI", "ba"], ooxml: doc([exact("a")]), twips: 400, expectedExact: 1 });
    expect(result).toEqual({ ok: false, problem: "nội dung đoạn 2 bị thay đổi" });
  });

  it("không đạt khi Exactly không được ghi nhận đủ", () => {
    const result = verifyExactPass({ before, after: before, ooxml: doc([exact("một"), auto("hai"), auto("ba")]), twips: 400, expectedExact: 3 });
    expect(result).toEqual({ ok: false, problem: "chỉ 1 trên 3 đoạn có giãn dòng Exactly" });
  });
});
