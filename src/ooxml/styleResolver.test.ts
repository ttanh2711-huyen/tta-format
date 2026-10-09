import { readParagraph, type ParagraphModel } from "../rules/docModel";
import { parseStyleTable, resolveParagraph, styleNameOf } from "./styleResolver";

const styles = `<w:styles>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="ND30NoiDung"><w:name w:val="ND30_NoiDung"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="120" w:after="120" w:line="400" w:lineRule="exact"/><w:ind w:firstLine="567"/><w:jc w:val="both"/></w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Con"><w:name w:val="Con"/><w:basedOn w:val="ND30NoiDung"/><w:pPr><w:spacing w:line="280" w:lineRule="auto"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>
</w:styles>`;

const table = parseStyleTable(styles);
const para = (xml: string): ParagraphModel => readParagraph(xml);

describe("parseStyleTable", () => {
  it("đọc tên style theo mã", () => {
    expect(styleNameOf("ND30NoiDung", table)).toBe("ND30_NoiDung");
    expect(styleNameOf("Normal", table)).toBe("Normal");
    expect(styleNameOf("khong-co", table)).toBeUndefined();
  });

  it("gói rỗng hoặc không có phần styles thì vẫn dùng được", () => {
    const empty = parseStyleTable("");
    expect(resolveParagraph(para("<w:p><w:r><w:t>a</w:t></w:r></w:p>"), empty).text).toBe("a");
  });
});

describe("resolveParagraph: giá trị hiệu lực", () => {
  it("đoạn không khai báo style dùng style mặc định và giá trị mặc định của tài liệu", () => {
    const resolved = resolveParagraph(para("<w:p><w:r><w:t>Xin chào</w:t></w:r></w:p>"), table);
    expect(resolved).toMatchObject({ fontName: "Calibri", sizePt: 11, spaceAfterPt: 8 });
    expect(resolved.lineMultiple).toBeCloseTo(1.08, 2);
  });

  it("đoạn dùng style của bộ chuẩn: lấy căn lề, thụt, giãn dòng Exactly và phông từ style", () => {
    const resolved = resolveParagraph(para('<w:p><w:pPr><w:pStyle w:val="ND30NoiDung"/></w:pPr><w:r><w:t>Nội dung</w:t></w:r></w:p>'), table);
    expect(resolved).toMatchObject({ align: "both", firstLineIndentMm: 10, spaceBeforePt: 6, spaceAfterPt: 6, lineExactPt: 20, fontName: "Times New Roman", sizePt: 14 });
    expect(resolved.lineMultiple).toBeUndefined();
  });

  it("định dạng trực tiếp trên đoạn thắng style", () => {
    const resolved = resolveParagraph(para('<w:p><w:pPr><w:pStyle w:val="ND30NoiDung"/><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:sz w:val="26"/></w:rPr><w:t>x</w:t></w:r></w:p>'), table);
    expect(resolved).toMatchObject({ align: "left", sizePt: 13, fontName: "Times New Roman" });
  });

  it("style kế thừa: giãn dòng của style con thay quy tắc của style cha, các giá trị khác giữ từ cha", () => {
    const resolved = resolveParagraph(para('<w:p><w:pPr><w:pStyle w:val="Con"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'), table);
    expect(resolved.lineExactPt).toBeUndefined();
    expect(resolved.lineMultiple).toBeCloseTo(1.17, 2);
    expect(resolved).toMatchObject({ bold: true, align: "both", spaceBeforePt: 6, sizePt: 14 });
  });

  it("vòng lặp basedOn không làm treo", () => {
    const loop = parseStyleTable('<w:styles><w:style w:type="paragraph" w:styleId="A"><w:name w:val="A"/><w:basedOn w:val="B"/></w:style><w:style w:type="paragraph" w:styleId="B"><w:name w:val="B"/><w:basedOn w:val="A"/></w:style></w:styles>');
    expect(resolveParagraph(para('<w:p><w:pPr><w:pStyle w:val="A"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'), loop).text).toBe("x");
  });

  it("không đổi chữ và các đặc điểm riêng của đoạn", () => {
    const resolved = resolveParagraph(para('<w:p><w:pPr><w:pStyle w:val="ND30NoiDung"/><w:pBdr><w:bottom w:val="single"/></w:pBdr></w:pPr><w:r><w:t>Chữ</w:t></w:r></w:p>'), table);
    expect(resolved).toMatchObject({ text: "Chữ", styleId: "ND30NoiDung", borderBottom: true });
  });
});
