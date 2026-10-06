import {
  bodyPart,
  hasBottomBorder,
  hasExactLineRule,
  hasPageField,
  paragraphStyleIds,
  readJustification,
  readLineRule,
  readParagraphIndentTwips,
  readPgNumTypeStart,
  readTableIndentTwips,
  styleDefinition,
  styleIdByName,
  tableBordersAllNone,
} from "./ooxmlFacts";

describe("hasBottomBorder", () => {
  it("nhận ra viền dưới của đoạn", () => {
    const xml = `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>`;
    expect(hasBottomBorder(xml)).toBe(true);
  });

  it("không coi viền \"none\" hoặc \"nil\" là có viền", () => {
    expect(hasBottomBorder(`<w:pBdr><w:bottom w:val="none" w:sz="0"/></w:pBdr>`)).toBe(false);
    expect(hasBottomBorder(`<w:pBdr><w:bottom w:val="nil"/></w:pBdr>`)).toBe(false);
  });

  it("không nhầm viền trên với viền dưới và đoạn không có viền", () => {
    expect(hasBottomBorder(`<w:pBdr><w:top w:val="single"/></w:pBdr>`)).toBe(false);
    expect(hasBottomBorder(`<w:p><w:r><w:t>chữ</w:t></w:r></w:p>`)).toBe(false);
  });
});

describe("readParagraphIndentTwips", () => {
  it("đọc thụt trái và thụt phải", () => {
    expect(readParagraphIndentTwips(`<w:pPr><w:ind w:left="1500" w:right="1600"/></w:pPr>`)).toEqual({
      left: 1500,
      right: 1600,
    });
  });

  it("chấp nhận tên thuộc tính start, end của phiên bản chuẩn mới", () => {
    expect(readParagraphIndentTwips(`<w:ind w:start="300" w:end="400"/>`)).toEqual({ left: 300, right: 400 });
  });

  it("trả về đối tượng rỗng khi không có thụt dòng", () => {
    expect(readParagraphIndentTwips(`<w:pPr><w:jc w:val="center"/></w:pPr>`)).toEqual({});
  });
});

describe("tableBordersAllNone", () => {
  const none = (side: string) => `<w:${side} w:val="none" w:sz="4" w:space="0" w:color="auto"/>`;
  const sides = ["top", "left", "bottom", "right", "insideH", "insideV"];

  it("đúng khi cả 6 cạnh đều không viền", () => {
    const xml = `<w:tblPr><w:tblBorders>${sides.map(none).join("")}</w:tblBorders></w:tblPr>`;
    expect(tableBordersAllNone(xml)).toBe(true);
  });

  it("sai khi có một cạnh còn viền", () => {
    const withBorder = sides.map((s) => (s === "insideH" ? `<w:insideH w:val="single" w:sz="4"/>` : none(s)));
    expect(tableBordersAllNone(`<w:tblBorders>${withBorder.join("")}</w:tblBorders>`)).toBe(false);
  });

  it("sai khi bảng không khai báo viền rõ ràng", () => {
    expect(tableBordersAllNone(`<w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>`)).toBe(false);
  });
});

describe("readTableIndentTwips", () => {
  it("đọc thụt lề bảng, kể cả giá trị âm như trong file mẫu", () => {
    expect(readTableIndentTwips(`<w:tblPr><w:tblInd w:w="-800" w:type="dxa"/></w:tblPr>`)).toBe(-800);
  });

  it("không có thì trả về undefined", () => {
    expect(readTableIndentTwips(`<w:tblPr></w:tblPr>`)).toBeUndefined();
  });
});

describe("hasExactLineRule", () => {
  const exact = `<w:spacing w:before="120" w:after="120" w:line="400" w:lineRule="exact"/>`;

  it("nhận ra giãn dòng Exactly, đúng giá trị twips nếu có yêu cầu", () => {
    expect(hasExactLineRule(exact)).toBe(true);
    expect(hasExactLineRule(exact, 400)).toBe(true);
    expect(hasExactLineRule(exact, 360)).toBe(false);
  });

  it("giãn dòng bội số (auto) không phải Exactly", () => {
    expect(hasExactLineRule(`<w:spacing w:line="280" w:lineRule="auto"/>`)).toBe(false);
  });

  it("chấp nhận thứ tự thuộc tính khác", () => {
    expect(hasExactLineRule(`<w:spacing w:lineRule="exact" w:line="440"/>`, 440)).toBe(true);
  });
});

describe("hasPageField", () => {
  it("nhận ra trường PAGE dạng fldSimple và dạng instrText", () => {
    expect(hasPageField(`<w:fldSimple w:instr=" PAGE   \\* MERGEFORMAT "><w:r><w:t>1</w:t></w:r></w:fldSimple>`)).toBe(true);
    expect(hasPageField(`<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>`)).toBe(true);
  });

  it("không nhầm với NUMPAGES hoặc chữ thường", () => {
    expect(hasPageField(`<w:fldSimple w:instr=" NUMPAGES "/>`)).toBe(false);
    expect(hasPageField(`<w:t>trang 1</w:t>`)).toBe(false);
  });
});

describe("readPgNumTypeStart", () => {
  it("đọc số trang bắt đầu của phần", () => {
    expect(readPgNumTypeStart(`<w:sectPr><w:pgNumType w:start="1"/></w:sectPr>`)).toBe(1);
  });

  it("không có thì trả về undefined", () => {
    expect(readPgNumTypeStart(`<w:sectPr><w:pgSz w:w="11906"/></w:sectPr>`)).toBeUndefined();
  });
});

describe("readJustification", () => {
  it("đọc kiểu căn lề đầu tiên", () => {
    expect(readJustification(`<w:pPr><w:jc w:val="center"/></w:pPr>`)).toBe("center");
    expect(readJustification(`<w:pPr></w:pPr>`)).toBeUndefined();
  });
});

describe("bodyPart", () => {
  const pkg =
    `<pkg:package><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>` +
    `<w:p><w:r><w:t>nội dung</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles>` +
    `<w:style w:styleId="Cu"><w:pPr><w:spacing w:line="400" w:lineRule="exact"/></w:pPr></w:style>` +
    `</w:styles></pkg:xmlData></pkg:part></pkg:package>`;

  it("chỉ lấy phần thân tài liệu, không lấy phần styles", () => {
    expect(bodyPart(pkg)).toContain("nội dung");
    expect(bodyPart(pkg)).not.toContain("w:styleId");
  });

  it("nhờ đó không nhầm giãn dòng Exactly của style khác với giãn dòng của đoạn", () => {
    expect(hasExactLineRule(pkg, 400)).toBe(true);
    expect(hasExactLineRule(bodyPart(pkg), 400)).toBe(false);
  });

  it("không có thân tài liệu thì trả về chuỗi rỗng", () => {
    expect(bodyPart("<pkg:package/>")).toBe("");
  });
});

describe("styleDefinition", () => {
  const styles =
    `<w:style w:type="paragraph" w:styleId="A1"><w:name w:val="A"/><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:style>` +
    `<w:style w:type="paragraph" w:customStyle="1" w:styleId="B2"><w:name w:val="B"/><w:pPr><w:spacing w:line="400" w:lineRule="exact"/></w:pPr></w:style>`;

  it("lấy đúng định nghĩa của style theo mã", () => {
    expect(hasExactLineRule(styleDefinition(styles, "B2"), 400)).toBe(true);
    expect(hasExactLineRule(styleDefinition(styles, "A1"))).toBe(false);
  });

  it("style không tồn tại thì trả về chuỗi rỗng", () => {
    expect(styleDefinition(styles, "KhongCo")).toBe("");
  });
});

describe("readLineRule", () => {
  it("đọc quy tắc giãn dòng đầu tiên", () => {
    expect(readLineRule(`<w:spacing w:line="400" w:lineRule="exact"/>`)).toBe("exact");
    expect(readLineRule(`<w:spacing w:line="280" w:lineRule="auto"/>`)).toBe("auto");
    expect(readLineRule(`<w:spacing w:before="120"/>`)).toBeUndefined();
  });
});

describe("paragraphStyleIds", () => {
  it("liệt kê mã style của các đoạn", () => {
    const xml = `<w:p><w:pPr><w:pStyle w:val="A1"/></w:pPr></w:p><w:p><w:pPr><w:pStyle w:val="B2"/></w:pPr></w:p>`;
    expect(paragraphStyleIds(xml)).toEqual(["A1", "B2"]);
    expect(paragraphStyleIds(`<w:p/>`)).toEqual([]);
  });
});

describe("styleIdByName", () => {
  const styles =
    `<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>` +
    `<w:style w:type="paragraph" w:customStyle="1" w:styleId="ND30C1"><w:name w:val="ND30_C1"/></w:style>`;

  it("Word bỏ dấu gạch dưới khi tạo mã style từ tên; tra mã theo tên đúng", () => {
    expect(styleIdByName(styles, "ND30_C1")).toBe("ND30C1");
    expect(styleIdByName(styles, "Normal")).toBe("Normal");
  });

  it("tên không tồn tại thì trả về undefined", () => {
    expect(styleIdByName(styles, "KhongCo")).toBeUndefined();
  });

  it("tên có ký tự đặc biệt của biểu thức chính quy vẫn tra đúng", () => {
    const special = `<w:style w:styleId="X1"><w:name w:val="A (b) + c"/></w:style>`;
    expect(styleIdByName(special, "A (b) + c")).toBe("X1");
  });
});
