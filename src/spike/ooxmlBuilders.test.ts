import {
  exactParagraphBody,
  flatOpc,
  headerTableBody,
  lineParagraphBody,
  mmToTwips,
  pageNumberParagraphBody,
  sectionBreakRestartBody,
  styledParagraphBody,
  styleDefinitionXml,
  styleImportJson,
  withPgNumTypeStart,
  withTrailingParagraph,
} from "./ooxmlBuilders";
import {
  hasBottomBorder,
  hasExactLineRule,
  hasPageField,
  readJustification,
  readParagraphIndentTwips,
  readPgNumTypeStart,
  readTableIndentTwips,
  tableBordersAllNone,
} from "../ooxml/facts";

/** Kiểm tra XML đóng mở thẻ cân đối, đủ để phát hiện lỗi dựng chuỗi. */
function isBalanced(xml: string): boolean {
  const stack: string[] = [];
  const tag = /<(\/?)([A-Za-z][\w:.-]*)([^>]*?)(\/?)>/g;
  let match: RegExpExecArray | null;
  while ((match = tag.exec(xml)) !== null) {
    const [, closing, name, , selfClosing] = match;
    if (selfClosing) continue;
    if (closing) {
      if (stack.pop() !== name) return false;
    } else {
      stack.push(name);
    }
  }
  return stack.length === 0;
}

describe("mmToTwips", () => {
  it("1 inch = 25,4 mm = 1440 twips; vùng chữ 160 mm = 9071 twips", () => {
    expect(mmToTwips(25.4)).toBe(1440);
    expect(mmToTwips(160)).toBe(9071);
    expect(mmToTwips(30)).toBe(1701);
  });
});

describe("lineParagraphBody", () => {
  const xml = lineParagraphBody({ leftTwips: 1500, rightTwips: 1600 });

  it("có viền dưới, thụt hai bên theo yêu cầu và canh giữa", () => {
    expect(hasBottomBorder(xml)).toBe(true);
    expect(readParagraphIndentTwips(xml)).toEqual({ left: 1500, right: 1600 });
    expect(readJustification(xml)).toBe("center");
  });

  it("XML cân đối", () => {
    expect(isBalanced(xml)).toBe(true);
  });
});

describe("headerTableBody", () => {
  it("bảng 3 hàng, 2 cột, không viền, vừa vùng chữ", () => {
    const xml = headerTableBody({ totalWidthTwips: 9071, leftRatio: 0.45, indentTwips: 0 });
    expect(tableBordersAllNone(xml)).toBe(true);
    expect(readTableIndentTwips(xml)).toBe(0);
    expect(xml.match(/<w:tr[ >]/g)).toHaveLength(3);
    expect(xml.match(/<w:gridCol /g)).toHaveLength(2);
    expect(xml).toContain('w:w="4082"');
    expect(xml).toContain('w:w="4989"');
    expect(isBalanced(xml)).toBe(true);
  });

  it("bảng rộng hơn vùng chữ và thụt âm như file mẫu cũ", () => {
    const xml = headerTableBody({ totalWidthTwips: 10400, leftRatio: 4700 / 10400, indentTwips: -800 });
    expect(readTableIndentTwips(xml)).toBe(-800);
    expect(xml).toContain('w:w="10400"');
    expect(isBalanced(xml)).toBe(true);
  });

  it("giữ nguyên chữ tiếng Việt có dấu", () => {
    const xml = headerTableBody({ totalWidthTwips: 9071, leftRatio: 0.45, indentTwips: 0 });
    expect(xml).toContain("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
    expect(xml).toContain("Độc lập - Tự do - Hạnh phúc");
  });
});

describe("styleDefinitionXml", () => {
  it("giãn dòng Exactly 20 pt = 400 twips, cỡ chữ 14 pt = 28 nửa điểm", () => {
    const xml = styleDefinitionXml({ id: "ND30Spike", name: "ND30_Spike", fontPt: 14, exactLineSpacingPt: 20 });
    expect(hasExactLineRule(xml, 400)).toBe(true);
    expect(xml).toContain('<w:sz w:val="28"/>');
    expect(xml).toContain('w:styleId="ND30Spike"');
    expect(xml).toContain('<w:name w:val="ND30_Spike"/>');
    expect(isBalanced(xml)).toBe(true);
  });

  it("không có giãn dòng Exactly khi không yêu cầu", () => {
    const xml = styleDefinitionXml({ id: "ND30Spike", name: "ND30_Spike", fontPt: 13 });
    expect(hasExactLineRule(xml)).toBe(false);
    expect(xml).toContain('<w:sz w:val="26"/>');
  });
});

describe("styleImportJson", () => {
  it("JSON hợp lệ, một kiểu đoạn, đúng cỡ chữ và giãn dòng", () => {
    const parsed = JSON.parse(styleImportJson({ name: "ND30_Spike", fontPt: 14, lineSpacingPt: 20 }));
    expect(parsed.styles).toHaveLength(1);
    expect(parsed.styles[0]).toMatchObject({
      nameLocal: "ND30_Spike",
      type: "Paragraph",
      baseStyle: "Normal",
      builtIn: false,
      font: { name: "Times New Roman", size: 14 },
      paragraphFormat: { lineSpacing: 20 },
    });
  });
});

describe("pageNumberParagraphBody", () => {
  it("đoạn canh giữa chứa trường PAGE", () => {
    const xml = pageNumberParagraphBody(14);
    expect(hasPageField(xml)).toBe(true);
    expect(readJustification(xml)).toBe("center");
    expect(xml).toContain('<w:sz w:val="28"/>');
    expect(isBalanced(xml)).toBe(true);
  });
});

describe("sectionBreakRestartBody", () => {
  it("ngắt phần có đặt số trang bắt đầu lại từ 1", () => {
    const xml = sectionBreakRestartBody();
    expect(readPgNumTypeStart(xml)).toBe(1);
    expect(isBalanced(xml)).toBe(true);
  });
});

describe("flatOpc", () => {
  it("gói OOXML phẳng có phần tài liệu và bọc đúng thân", () => {
    const xml = flatOpc({ bodyXml: "<w:p><w:r><w:t>Xin chào</w:t></w:r></w:p>" });
    expect(xml).toContain("<pkg:package");
    expect(xml).toContain('pkg:name="/word/document.xml"');
    expect(xml).toContain("Xin chào");
    expect(xml).not.toContain("/word/styles.xml");
    expect(isBalanced(xml)).toBe(true);
  });

  it("có phần styles và quan hệ tới styles khi truyền định nghĩa style", () => {
    const style = styleDefinitionXml({ id: "ND30Spike", name: "ND30_Spike", fontPt: 14 });
    const xml = flatOpc({ bodyXml: "<w:p/>", stylesXml: style });
    expect(xml).toContain('pkg:name="/word/styles.xml"');
    expect(xml).toContain("relationships/styles");
    expect(xml).toContain('w:styleId="ND30Spike"');
    expect(isBalanced(xml)).toBe(true);
  });
});

describe("withTrailingParagraph", () => {
  it("thêm một đoạn rỗng phía sau để đoạn trước giữ được thuộc tính khi chèn", () => {
    const body = lineParagraphBody({ leftTwips: 100, rightTwips: 100 });
    const wrapped = withTrailingParagraph(body);
    expect(wrapped.startsWith(body)).toBe(true);
    expect(wrapped.slice(body.length)).toBe("<w:p/>");
    expect(isBalanced(wrapped)).toBe(true);
  });
});

describe("styledParagraphBody", () => {
  it("đoạn dùng style theo mã và giữ chữ tiếng Việt", () => {
    const xml = styledParagraphBody("NDThuB", "Thử style tiếng Việt");
    expect(xml).toContain('<w:pStyle w:val="NDThuB"/>');
    expect(xml).toContain("Thử style tiếng Việt");
    expect(isBalanced(xml)).toBe(true);
  });
});

describe("exactParagraphBody", () => {
  it("đoạn có giãn dòng Exactly đặt trực tiếp, 20 pt = 400 twips", () => {
    const xml = exactParagraphBody("Giãn dòng Exactly", 20);
    expect(hasExactLineRule(xml, 400)).toBe(true);
    expect(isBalanced(xml)).toBe(true);
  });
});

describe("withPgNumTypeStart", () => {
  const sect = (extra = "") => `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:left="1701"/>${extra}<w:cols w:space="708"/></w:sectPr>`;
  const doc = `<w:p><w:pPr>${sect()}</w:pPr></w:p><w:p><w:pPr>${sect()}</w:pPr></w:p>${sect()}`;

  it("thêm pgNumType vào đúng phần thứ n, các phần khác giữ nguyên", () => {
    const result = withPgNumTypeStart(doc, 1, 1);
    const sections = result.match(/<w:sectPr>[\s\S]*?<\/w:sectPr>/g) ?? [];
    expect(sections).toHaveLength(3);
    expect(sections[0]).not.toContain("pgNumType");
    expect(sections[1]).toContain('<w:pgNumType w:start="1"/>');
    expect(sections[2]).not.toContain("pgNumType");
    expect(sections[1].indexOf("pgMar")).toBeLessThan(sections[1].indexOf("pgNumType"));
    expect(sections[1].indexOf("pgNumType")).toBeLessThan(sections[1].indexOf("<w:cols"));
    expect(isBalanced(result)).toBe(true);
  });

  it("thay giá trị khi phần đã có pgNumType", () => {
    const withExisting = `<w:p><w:pPr>${sect('<w:pgNumType w:start="5"/>')}</w:pPr></w:p>`;
    const result = withPgNumTypeStart(withExisting, 0, 1);
    expect(readPgNumTypeStart(result)).toBe(1);
    expect(result.match(/pgNumType/g)).toHaveLength(1);
  });

  it("chỉ số phần không tồn tại là lỗi", () => {
    expect(() => withPgNumTypeStart(doc, 7, 1)).toThrow(RangeError);
  });
});
