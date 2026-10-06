import { hasExactLineRule, readLineRule } from "./facts";
import { prepareParagraphOoxml, stripBodySectPr, withExactLineSpacing } from "./exactSpacing";

const wrap = (body: string, styles = "") =>
  `<pkg:package><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${body}</w:body></w:document></pkg:xmlData></pkg:part>${styles}</pkg:package>`;

describe("withExactLineSpacing", () => {
  it("đoạn chưa có pPr: thêm pPr chứa giãn dòng Exactly, 20 pt = 400 twips", () => {
    const result = withExactLineSpacing(wrap("<w:p><w:r><w:t>Xin chào</w:t></w:r></w:p>"), 20);
    expect(result).toContain('<w:p><w:pPr><w:spacing w:line="400" w:lineRule="exact"/></w:pPr><w:r>');
    expect(result).toContain("Xin chào");
  });

  it("đoạn có pPr: chèn giãn dòng sau pStyle và trước jc theo đúng thứ tự của lược đồ", () => {
    const result = withExactLineSpacing(wrap(`<w:p><w:pPr><w:pStyle w:val="A"/><w:jc w:val="both"/></w:pPr></w:p>`), 20);
    expect(result).toContain('<w:pPr><w:pStyle w:val="A"/><w:spacing w:line="400" w:lineRule="exact"/><w:jc w:val="both"/></w:pPr>');
  });

  it("giữ nguyên giãn trước, sau đã có và chỉ đổi line, lineRule", () => {
    const result = withExactLineSpacing(wrap(`<w:p><w:pPr><w:spacing w:before="120" w:after="120" w:line="280" w:lineRule="auto"/></w:pPr></w:p>`), 22);
    expect(result).toContain('<w:spacing w:before="120" w:after="120" w:line="440" w:lineRule="exact"/>');
    expect(result).not.toContain('w:lineRule="auto"');
  });

  it("thẻ spacing chưa có line thì bổ sung", () => {
    const result = withExactLineSpacing(wrap(`<w:p><w:pPr><w:spacing w:before="120"/></w:pPr></w:p>`), 18);
    expect(result).toContain('<w:spacing w:before="120" w:line="360" w:lineRule="exact"/>');
  });

  it("chèn trước ind, rPr hoặc sectPr khi đoạn có các phần tử này", () => {
    expect(withExactLineSpacing(wrap(`<w:p><w:pPr><w:ind w:firstLine="567"/></w:pPr></w:p>`), 20)).toContain(
      '<w:spacing w:line="400" w:lineRule="exact"/><w:ind w:firstLine="567"/>',
    );
    expect(withExactLineSpacing(wrap(`<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr></w:p>`), 20)).toContain(
      '<w:spacing w:line="400" w:lineRule="exact"/><w:rPr><w:b/></w:rPr>',
    );
    expect(withExactLineSpacing(wrap(`<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906"/></w:sectPr></w:pPr></w:p>`), 20)).toContain(
      '<w:spacing w:line="400" w:lineRule="exact"/><w:sectPr>',
    );
  });

  it("đoạn rỗng tự đóng và pPr tự đóng cũng được xử lý", () => {
    expect(withExactLineSpacing(wrap("<w:p/>"), 20)).toContain('<w:p><w:pPr><w:spacing w:line="400" w:lineRule="exact"/></w:pPr></w:p>');
    expect(withExactLineSpacing(wrap("<w:p><w:pPr/></w:p>"), 20)).toContain('<w:pPr><w:spacing w:line="400" w:lineRule="exact"/></w:pPr>');
  });

  it("xử lý mọi đoạn trong thân, nhận thuộc tính của thẻ w:p", () => {
    const result = withExactLineSpacing(wrap(`<w:p w:rsidR="00A1"><w:r><w:t>Một</w:t></w:r></w:p><w:p><w:r><w:t>Hai</w:t></w:r></w:p>`), 20);
    expect(result.match(/w:lineRule="exact"/g)).toHaveLength(2);
    expect(result).toContain('<w:p w:rsidR="00A1"><w:pPr>');
  });

  it("không đụng phần styles của gói", () => {
    const styles = `<pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles><w:style w:styleId="X"><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:style></w:styles></pkg:xmlData></pkg:part>`;
    const result = withExactLineSpacing(wrap("<w:p><w:r><w:t>a</w:t></w:r></w:p>", styles), 20);
    expect(result).toContain('<w:spacing w:line="240" w:lineRule="auto"/>');
    expect(hasExactLineRule(result, 400)).toBe(true);
  });

  it("chạy lại nhiều lần cho cùng kết quả", () => {
    const once = withExactLineSpacing(wrap(`<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>`), 20);
    expect(withExactLineSpacing(once, 20)).toBe(once);
    expect(readLineRule(once)).toBe("exact");
  });

  it("không có thân tài liệu thì trả về nguyên bản", () => {
    expect(withExactLineSpacing("<pkg:package/>", 20)).toBe("<pkg:package/>");
  });
});

describe("stripBodySectPr", () => {
  const sect = `<w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`;

  it("bỏ thuộc tính trang ở cấp thân (gói OOXML của một đoạn có kèm), không nên chèn lại", () => {
    const result = stripBodySectPr(wrap(`<w:p><w:r><w:t>a</w:t></w:r></w:p>${sect}`));
    expect(result).not.toContain("sectPr");
    expect(result).toContain("<w:t>a</w:t>");
  });

  it("giữ sectPr nằm trong đoạn (ngắt phần thật)", () => {
    const body = `<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906"/></w:sectPr></w:pPr></w:p>${sect}`;
    const result = stripBodySectPr(wrap(body));
    expect(result).toContain("<w:pPr><w:sectPr><w:pgSz");
    expect(result.match(/<w:sectPr>/g)).toHaveLength(1);
  });

  it("không có sectPr cấp thân thì trả về nguyên bản", () => {
    const xml = wrap("<w:p/>");
    expect(stripBodySectPr(xml)).toBe(xml);
  });
});

describe("prepareParagraphOoxml", () => {
  it("bỏ sectPr cấp thân rồi đặt Exactly", () => {
    const xml = wrap(`<w:p><w:r><w:t>a</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906"/></w:sectPr>`);
    const result = prepareParagraphOoxml(xml, 20);
    expect(result).not.toContain("sectPr");
    expect(hasExactLineRule(result, 400)).toBe(true);
  });
});
