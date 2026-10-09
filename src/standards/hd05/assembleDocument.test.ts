import { expectMatchesGolden } from "../../ooxml/__golden__/golden";
import { BODY_MARKER } from "../../ooxml/documentAssembler";
import { extractDocModel, type DocModel, type ParagraphModel, type TableModel } from "../../rules/docModel";
import { assemblePartyDocumentBody, type PartyAssembleInput } from "./assembleDocument";
import { partySubjectRow } from "./titleBlock";

const sizes = { theLoai: 16, trichYeu: 14, trichYeuCongVan: 12, kinhGui: 14 };

const letter: PartyAssembleInput = {
  header: {
    parentLines: ["Tỉnh ủy Đồng Tháp"],
    issuerLines: ["Ban Tổ chức"],
    sizes: { title: 15, parent: 14, issuer: 14, so: 14, placeDate: 14 },
    numberRow: { so: "Số 05-CV/BTCTU", placeDate: "Đồng Tháp, ngày 03 tháng 02 năm 2026" },
    subjectRow: { text: "về việc cử cán bộ dự hội nghị", sizePt: 12 },
    recipientsRow: { label: "Kính gửi:", italicLabel: true, lines: ["- Ban Thường vụ Tỉnh ủy."], sizePt: 14 },
    textWidthMm: 165,
  },
  title: { variant: "congVan", subject: "cử cán bộ dự hội nghị", recipients: ["Ban Thường vụ Tỉnh ủy"], basis: [], sizes, lineExactPt: 20, textWidthMm: 165 },
  footer: {
    recipients: ["Như trên"],
    luu: "VT, TCTU",
    drafter: "",
    authority: "K/T",
    positionLines: ["Trưởng ban", "Phó Trưởng ban"],
    signerName: "Nguyễn Văn A",
    confirmation: "",
    signGapLines: 6,
    sizes: { noiNhanTu: 14, noiNhanDanhSach: 12, quyenHan: 14, chucVu: 14, hoTen: 14, kyHieuNguoiSoan: 11 },
    textWidthMm: 165,
  },
};

function modelOf(input: PartyAssembleInput, body = {}): DocModel {
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${assemblePartyDocumentBody(input, body)}</w:body></w:document>`;
  return extractDocModel({ documentXml: xml });
}

describe("assemblePartyDocumentBody", () => {
  it("công văn: khối đầu (đã gồm về việc và Kính gửi), chỗ nội dung, khối cuối, đoạn đệm", () => {
    const blocks = modelOf(letter).blocks;
    expect(blocks.map((b) => b.kind)).toEqual(["table", "paragraph", "table", "paragraph"]);
    expect((blocks[1] as ParagraphModel).text).toBe(BODY_MARKER);
    expect((blocks[0] as TableModel).rows).toHaveLength(4);
  });

  it("chỗ đánh dấu nội dung xuất hiện đúng một lần", () => {
    expect(assemblePartyDocumentBody(letter).split(BODY_MARKER)).toHaveLength(2);
  });

  it("có thể loại và căn cứ: thể loại, trích yếu, 5 dấu gạch nối, căn cứ rồi chỗ nội dung", () => {
    const named: PartyAssembleInput = {
      ...letter,
      header: { ...letter.header, subjectRow: undefined, recipientsRow: undefined },
      title: { ...letter.title!, variant: "coTenLoai", typeName: "Quyết định", subject: "Về việc ban hành Quy chế", recipients: [], basis: ["Căn cứ Điều lệ Đảng", "Căn cứ Quy định 399-QĐ/TW"] },
    };
    const texts = modelOf(named).blocks.slice(1, 7).map((b) => (b as ParagraphModel).text);
    expect(texts).toEqual(["QUYẾT ĐỊNH", "Về việc ban hành Quy chế", "-----", "- Căn cứ Điều lệ Đảng;", "- Căn cứ Quy định 399-QĐ/TW.", BODY_MARKER]);
  });

  it("không có phần thể loại thì chèn một đoạn trống để hai bảng không dính nhau", () => {
    const kinds = modelOf({ ...letter, title: undefined, header: { ...letter.header, subjectRow: undefined, recipientsRow: undefined } }).blocks.map((b) => b.kind);
    expect(kinds).toEqual(["table", "paragraph", "paragraph", "table", "paragraph"]);
  });

  it("đoạn chỗ nội dung gắn style và Exactly nếu được yêu cầu", () => {
    const xml = assemblePartyDocumentBody(letter, { styleId: "HD05NoiDung", exactPt: 20 });
    expect(xml).toContain('<w:pStyle w:val="HD05NoiDung"/><w:spacing w:line="400" w:lineRule="exact"/>');
    const marker = modelOf(letter, { styleId: "HD05NoiDung", exactPt: 20 }).blocks[1] as ParagraphModel;
    expect(marker.lineExactPt).toBe(20);
  });

  it("dòng về việc lấy từ bộ dựng thể loại", () => {
    expect(partySubjectRow(letter.title!)).toEqual({ text: "về việc cử cán bộ dự hội nghị", sizePt: 12 });
  });

  it("golden: công văn", () => {
    expectMatchesGolden("document-party-letter.xml", assemblePartyDocumentBody(letter, { styleId: "HD05NoiDung", exactPt: 20 }).replace(/></g, ">\n<"));
  });
});
