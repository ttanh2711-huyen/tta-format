import { anonymizeModel, extractDocModel, type ParagraphModel, type TableModel } from "./docModel";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const RFONTS = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>';

function para(text: string, pPr = "", rPr = "") {
  const run = text === "" ? "" : `<w:r><w:rPr>${RFONTS}${rPr}</w:rPr><w:t>${text}</w:t></w:r>`;
  return `<w:p><w:pPr>${pPr}</w:pPr>${run}</w:p>`;
}

function doc(body: string, sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1701" w:header="709" w:footer="300"/></w:sectPr>') {
  return `<w:document ${W}><w:body>${body}${sectPr}</w:body></w:document>`;
}

const paragraphAt = (model: ReturnType<typeof extractDocModel>, index: number) => model.blocks[index] as ParagraphModel;

describe("extractDocModel: trang", () => {
  it("đổi twips sang mm cho khổ giấy, lề và khoảng cách header, footer", () => {
    const { page } = extractDocModel({ documentXml: doc("") });
    expect(page).toMatchObject({
      widthMm: 210,
      heightMm: 297,
      marginsMm: { top: 20, bottom: 20, left: 30, right: 20 },
      headerDistanceMm: 12.5,
      footerDistanceMm: 5.3,
      titlePg: false,
    });
  });

  it("nhận biết trang đầu khác biệt và trường số trang ở header", () => {
    const sect = '<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1701" w:header="709" w:footer="300"/><w:titlePg/></w:sectPr>';
    const header = '<w:hdr><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r></w:p></w:hdr>';
    const { page } = extractDocModel({ documentXml: doc("", sect), headerXmls: [header] });
    expect(page.titlePg).toBe(true);
    expect(page.pageNumberInHeader).toBe(true);
  });
});

describe("extractDocModel: đoạn", () => {
  it("đọc căn lề, giãn, giãn dòng, thụt dòng đầu và thuộc tính chữ của đoạn", () => {
    const xml = doc(
      para(
        "Nội dung thử",
        '<w:spacing w:before="120" w:after="120" w:line="280" w:lineRule="auto"/><w:ind w:firstLine="720"/><w:jc w:val="both"/>',
        '<w:b/><w:i/><w:sz w:val="28"/>',
      ),
    );
    expect(paragraphAt(extractDocModel({ documentXml: xml }), 0)).toEqual({
      kind: "paragraph",
      text: "Nội dung thử",
      align: "both",
      spaceBeforePt: 6,
      spaceAfterPt: 6,
      lineMultiple: 1.17,
      firstLineIndentMm: 12.7,
      fontName: "Times New Roman",
      sizePt: 14,
      bold: true,
      italic: true,
    });
  });

  it("giãn dòng Exactly và tối thiểu", () => {
    const xml = doc(para("a", '<w:spacing w:line="400" w:lineRule="exact"/>') + para("b", '<w:spacing w:line="300" w:lineRule="atLeast"/>'));
    const model = extractDocModel({ documentXml: xml });
    expect(paragraphAt(model, 0).lineExactPt).toBe(20);
    expect(paragraphAt(model, 1).lineAtLeastPt).toBe(15);
  });

  it("đoạn rỗng lấy thuộc tính chữ từ dấu đoạn", () => {
    const xml = doc('<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="26"/></w:rPr></w:pPr></w:p>');
    expect(paragraphAt(extractDocModel({ documentXml: xml }), 0)).toMatchObject({ text: "", fontName: "Times New Roman", sizePt: 13 });
  });

  it("giải mã ký tự đặc biệt của XML và bỏ chữ trong phần dự phòng của hình", () => {
    const xml = doc('<w:p><w:r><w:t>A &amp; B</w:t></w:r><mc:AlternateContent><mc:Fallback><w:r><w:t>BỎ</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>');
    expect(paragraphAt(extractDocModel({ documentXml: xml }), 0).text).toBe("A & B");
  });

  it("đường kẻ nổi trong đoạn được ghi lại kèm độ dài mm", () => {
    const anchor =
      '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="628650" cy="0"/><a:graphic><a:graphicData><wps:wsp><wps:spPr><a:prstGeom prst="line"/></wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent>';
    const xml = doc(`<w:p><w:r>${anchor}</w:r><w:r><w:t>CƠ QUAN</w:t></w:r></w:p>`);
    expect(paragraphAt(extractDocModel({ documentXml: xml }), 0)).toMatchObject({ text: "CƠ QUAN", floatingLinesMm: [17.5] });
  });

  it("viền dưới của đoạn", () => {
    const xml = doc(para("a", '<w:pBdr><w:bottom w:val="single" w:sz="4"/></w:pBdr>'));
    expect(paragraphAt(extractDocModel({ documentXml: xml }), 0).borderBottom).toBe(true);
  });
});

describe("extractDocModel: bảng", () => {
  const table = `<w:tbl><w:tblPr><w:tblW w:w="10400" w:type="dxa"/><w:tblInd w:w="-800" w:type="dxa"/><w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4700"/><w:gridCol w:w="5700"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4700" w:type="dxa"/></w:tcPr>${para("Trái")}${para("Trái 2")}</w:tc><w:tc><w:tcPr/>${para("Phải")}</w:tc></w:tr></w:tbl>`;

  it("đọc độ rộng, thụt, cột, viền và đoạn trong từng ô", () => {
    const model = extractDocModel({ documentXml: doc(table + para("Sau bảng")) });
    expect(model.blocks).toHaveLength(2);
    const t = model.blocks[0] as TableModel;
    expect(t).toMatchObject({ kind: "table", widthMm: 183.4, indentMm: -14.1, columnsMm: [82.9, 100.5], bordersNone: true });
    expect(t.rows).toHaveLength(1);
    expect(t.rows[0]).toHaveLength(2);
    expect(t.rows[0][0].map((p) => p.text)).toEqual(["Trái", "Trái 2"]);
    expect(t.rows[0][1].map((p) => p.text)).toEqual(["Phải"]);
    expect(paragraphAt(model, 1).text).toBe("Sau bảng");
  });

  it("bảng không khai báo viền riêng không bị coi là không viền", () => {
    const t = '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>';
    expect((extractDocModel({ documentXml: doc(t) }).blocks[0] as TableModel).bordersNone).toBe(false);
  });

  it("bảng lồng nhau không làm lệch việc tách ô ngoài", () => {
    const inner = '<w:tbl><w:tblGrid><w:gridCol w:w="500"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Trong</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const outer = `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${inner}<w:p><w:r><w:t>Ngoài</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const t = extractDocModel({ documentXml: doc(outer) }).blocks[0] as TableModel;
    expect(t.rows).toHaveLength(1);
    expect(t.rows[0][0].map((p) => p.text)).toContain("Ngoài");
  });
});

describe("anonymizeModel: ẩn danh nội dung", () => {
  const model = extractDocModel({
    documentXml: doc(
      para("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "", "<w:b/>") +
        para("Độc lập - Tự do - Hạnh phúc") +
        para("Số:        /CCTL-ABC") +
        para("Kính gửi: Sở Nông nghiệp") +
        para("Nguyễn Văn A, 0909123456, a@example.com") +
        para("Nơi nhận:"),
    ),
  });

  it("giữ nguyên chuỗi công khai của thể thức", () => {
    const texts = anonymizeModel(model).blocks.map((b) => (b as ParagraphModel).text);
    expect(texts[0]).toBe("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
    expect(texts[1]).toBe("Độc lập - Tự do - Hạnh phúc");
    expect(texts[5]).toBe("Nơi nhận:");
  });

  it("chỉ giữ nhãn đầu dòng, thay phần còn lại bằng số ký tự", () => {
    const texts = anonymizeModel(model).blocks.map((b) => (b as ParagraphModel).text);
    expect(texts[2]).toBe("Số: […]");
    expect(texts[3]).toBe("Kính gửi: […]");
  });

  it("nội dung khác, kể cả họ tên, số điện thoại, email, bị thay hoàn toàn", () => {
    const json = JSON.stringify(anonymizeModel(model));
    expect(json).not.toMatch(/Nguyễn|0909|example\.com|CCTL|Sở Nông/);
    expect((anonymizeModel(model).blocks[4] as ParagraphModel).text).toBe("[…]");
  });

  it("giữ nguyên mọi thuộc tính định dạng", () => {
    const original = model.blocks[0] as ParagraphModel;
    const anonymous = anonymizeModel(model).blocks[0] as ParagraphModel;
    expect({ ...anonymous, text: "" }).toEqual({ ...original, text: "" });
  });
});
