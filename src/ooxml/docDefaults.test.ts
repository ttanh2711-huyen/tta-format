import { extractDocDefaults, flatOpc } from "./flatOpc";

const DEFAULTS = `<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:before="120" w:after="120"/><w:ind w:firstLine="720"/></w:pPr></w:pPrDefault></w:docDefaults>`;

describe("docDefaults trong gói OOXML", () => {
  it("lấy phần docDefaults từ OOXML của tài liệu", () => {
    const ooxml = `<pkg:package><pkg:part><pkg:xmlData><w:styles>${DEFAULTS}<w:style w:styleId="Normal"/></w:styles></pkg:xmlData></pkg:part></pkg:package>`;
    expect(extractDocDefaults(ooxml)).toBe(DEFAULTS);
    expect(extractDocDefaults("<w:styles/>")).toBeUndefined();
  });

  it("bỏ phần tử w14 vì gói không khai báo không gian tên đó (Word báo InvalidArgument)", () => {
    const withLigatures = `<w:docDefaults><w:rPrDefault><w:rPr><w:kern w:val="2"/><w14:ligatures w14:val="standardContextual"/></w:rPr></w:rPrDefault></w:docDefaults>`;
    expect(extractDocDefaults(`<w:styles>${withLigatures}</w:styles>`)).not.toContain("w14:");
  });

  it("gói có docDefaults của tài liệu đích để Word không bỏ thụt dòng, giãn đoạn bằng 0 khi chèn", () => {
    const xml = flatOpc({ bodyXml: "<w:p/>", docDefaultsXml: DEFAULTS });
    expect(xml).toContain(`<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${DEFAULTS}</w:styles>`);
    expect(xml).toContain('Target="styles.xml"');
  });

  it("không truyền thì gói không có phần styles như trước", () => {
    expect(flatOpc({ bodyXml: "<w:p/>" })).not.toContain("styles.xml");
  });

  it("docDefaults đứng trước các style", () => {
    const xml = flatOpc({ bodyXml: "<w:p/>", stylesXml: "<w:style/>", docDefaultsXml: DEFAULTS });
    expect(xml.indexOf("<w:docDefaults>")).toBeLessThan(xml.indexOf("<w:style/>"));
  });
});
