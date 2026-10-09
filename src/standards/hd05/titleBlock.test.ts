import { expectMatchesGolden } from "../../ooxml/__golden__/golden";
import { textWidthMm } from "../../ooxml/textWidth";
import { extractDocModel, type DocModel, type ParagraphModel } from "../../rules/docModel";
import { buildRegistry, defaultValues } from "../../settings/formatOptions";
import { HD05 } from "./profile";
import { buildPartyTitleBody, dashList, partySubjectRow, partyTitleFromOptions, validatePartyTitle, type PartyTitleInput } from "./titleBlock";

const base: PartyTitleInput = {
  variant: "coTenLoai",
  typeName: "Quyết định",
  subject: "Về việc ban hành Quy chế làm việc của Ban Chấp hành",
  recipients: [],
  basis: [],
  sizes: { theLoai: 16, trichYeu: 14, trichYeuCongVan: 12, kinhGui: 14 },
  lineExactPt: 20,
  textWidthMm: 165,
};

function blocks(input: PartyTitleInput): ParagraphModel[] {
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${buildPartyTitleBody(input)}</w:body></w:document>`;
  return (extractDocModel({ documentXml: xml }) as DocModel).blocks as ParagraphModel[];
}

describe("dashList: danh sách gạch đầu dòng", () => {
  it("gạch đầu dòng, ';' cuối dòng, '.' cuối dòng cuối", () => {
    expect(dashList(["Căn cứ A", "Căn cứ B", "Căn cứ C"])).toEqual(["- Căn cứ A;", "- Căn cứ B;", "- Căn cứ C."]);
  });

  it("chuẩn hoá dấu người dùng đã gõ, bỏ dòng trống, gọn khoảng trắng", () => {
    expect(dashList(["- Căn cứ A;", "  –  Căn cứ   B. ", "", "   "])).toEqual(["- Căn cứ A;", "- Căn cứ B."]);
  });

  it("một mục thì kết thúc bằng dấu chấm; rỗng thì không có dòng", () => {
    expect(dashList(["Chỉ một"])).toEqual(["- Chỉ một."]);
    expect(dashList(["", " "])).toEqual([]);
  });
});

describe("buildPartyTitleBody: có thể loại", () => {
  const [type, subject, dashes] = blocks(base);

  it("thể loại in hoa đậm 16 pt căn giữa; trích yếu in thường đậm 14 pt", () => {
    expect(type).toMatchObject({ text: "QUYẾT ĐỊNH", bold: true, sizePt: 16, align: "center" });
    expect(subject).toMatchObject({ text: "Về việc ban hành Quy chế làm việc của Ban Chấp hành", bold: true, sizePt: 14, align: "center" });
  });

  it("5 dấu gạch nối luôn là dòng riêng dưới trích yếu, căn giữa", () => {
    expect(dashes).toMatchObject({ text: "-----", align: "center" });
    expect(blocks(base)).toHaveLength(3);
  });

  it("trích yếu nhiều dòng (dài) vẫn là một đoạn căn giữa, dấu gạch nối vẫn là dòng riêng", () => {
    const long = `Về việc ${"thực hiện nhiệm vụ trọng tâm ".repeat(8)}`.trim();
    const result = blocks({ ...base, subject: long });
    expect(result).toHaveLength(3);
    expect(result[1]).toMatchObject({ text: long, align: "center" });
    expect(result[2].text).toBe("-----");
  });

  it("gọn khoảng trắng và thoát ký tự XML", () => {
    const result = blocks({ ...base, typeName: " nghị  quyết ", subject: "A & B" });
    expect(result[0].text).toBe("NGHỊ QUYẾT");
    expect(result[1].text).toBe("A & B");
  });

  it("golden: có thể loại", () => {
    expectMatchesGolden("party-title-named.xml", buildPartyTitleBody(base).replace(/></g, ">\n<"));
  });
});

describe("buildPartyTitleBody: công văn", () => {
  const letter: PartyTitleInput = { ...base, variant: "congVan", typeName: undefined, subject: "cử cán bộ dự hội nghị", recipients: ["Ban Thường vụ Tỉnh ủy", "Các đồng chí Ủy viên Ban Chấp hành"] };

  it("không ghi tên loại, không có 5 dấu gạch nối; có Kính gửi nghiêng 14 pt và danh sách nơi nhận", () => {
    const result = blocks(letter);
    expect(result.map((p) => p.text)).toEqual(["Kính gửi:", "- Ban Thường vụ Tỉnh ủy;", "- Các đồng chí Ủy viên Ban Chấp hành."]);
    expect(result[0]).toMatchObject({ italic: true, sizePt: 14, align: "center" });
    expect(result[1].align).toBe("left");
    expect(result[1].italic).toBeFalsy();
    expect(result[1].lineExactPt).toBe(20);
  });

  it("dòng 'về việc' nằm ở hàng của bảng đầu văn bản, 12 pt nghiêng", () => {
    expect(partySubjectRow(letter)).toEqual({ text: "về việc cử cán bộ dự hội nghị", sizePt: 12 });
    expect(partySubjectRow({ ...letter, subject: "V/v cử cán bộ" })).toEqual({ text: "về việc cử cán bộ", sizePt: 12 });
    expect(partySubjectRow({ ...letter, subject: "Về việc cử cán bộ" })).toEqual({ text: "về việc cử cán bộ", sizePt: 12 });
    expect(partySubjectRow(base)).toBeUndefined();
  });

  it("golden: công văn", () => {
    expectMatchesGolden("party-title-letter.xml", buildPartyTitleBody(letter).replace(/></g, ">\n<"));
  });
});

describe("buildPartyTitleBody: tờ trình và căn cứ", () => {
  const trinh: PartyTitleInput = { ...base, variant: "toTrinh", typeName: "Tờ trình", subject: "Về việc xin chủ trương", recipients: ["Ban Thường vụ Tỉnh ủy"] };

  it("tờ trình: thể loại, trích yếu, 5 dấu gạch nối rồi 'Kính trình:' và danh sách", () => {
    expect(blocks(trinh).map((p) => p.text)).toEqual(["TỜ TRÌNH", "Về việc xin chủ trương", "-----", "Kính trình:", "- Ban Thường vụ Tỉnh ủy."]);
  });

  it("căn cứ ban hành: mỗi căn cứ một dòng gạch đầu dòng, dấu câu tự chuẩn hoá, căn đều, Exactly", () => {
    const result = blocks({ ...base, basis: ["Căn cứ Điều lệ Đảng", "- Căn cứ Quy định 399-QĐ/TW;", "Căn cứ kết quả cuộc họp."] });
    const basis = result.slice(3);
    expect(basis.map((p) => p.text)).toEqual(["- Căn cứ Điều lệ Đảng;", "- Căn cứ Quy định 399-QĐ/TW;", "- Căn cứ kết quả cuộc họp."]);
    expect(basis[0]).toMatchObject({ align: "both", sizePt: 14, lineExactPt: 20 });
  });

  it("golden: tờ trình có căn cứ", () => {
    expectMatchesGolden("party-title-trinh.xml", buildPartyTitleBody({ ...trinh, basis: ["Căn cứ Điều lệ Đảng", "Căn cứ Quy định 399-QĐ/TW"] }).replace(/></g, ">\n<"));
  });
});

describe("validatePartyTitle", () => {
  it("hợp lệ", () => {
    expect(validatePartyTitle(base)).toEqual([]);
  });

  it("có tên loại: cần tên loại và trích yếu", () => {
    expect(validatePartyTitle({ ...base, typeName: " " })).toEqual(["Hãy nhập tên thể loại văn bản."]);
    expect(validatePartyTitle({ ...base, subject: "" })).toEqual(["Hãy nhập trích yếu nội dung."]);
  });

  it("công văn cần trích yếu và ít nhất một nơi nhận", () => {
    const letter: PartyTitleInput = { ...base, variant: "congVan", typeName: undefined, recipients: [] };
    expect(validatePartyTitle(letter)).toEqual(["Hãy nhập ít nhất một nơi nhận (Kính gửi)."]);
    expect(validatePartyTitle({ ...letter, recipients: ["A"], subject: " " })).toEqual(["Hãy nhập trích yếu nội dung."]);
  });

  it("tờ trình cần nơi nhận (Kính trình)", () => {
    expect(validatePartyTitle({ ...base, variant: "toTrinh", recipients: [] })).toEqual(["Hãy nhập ít nhất một nơi nhận (Kính trình)."]);
  });
});

describe("partyTitleFromOptions", () => {
  it("cỡ chữ theo Tuỳ chọn nhanh mặc định, giãn dòng Exactly và vùng chữ chuẩn Đảng", () => {
    const values = defaultValues(buildRegistry(HD05));
    expect(partyTitleFromOptions(HD05, values, { variant: "coTenLoai", typeName: "a", subject: "b", recipients: [], basis: [] })).toMatchObject({
      sizes: { theLoai: 16, trichYeu: 14, trichYeuCongVan: 12, kinhGui: 14 },
      lineExactPt: 20,
      textWidthMm: 165,
    });
  });

  it("độ dài dấu gạch nối là 5", () => {
    expect(textWidthMm("-----", 14, true)).toBeGreaterThan(0);
    expect(blocks(base)[2].text).toHaveLength(5);
  });
});
