import { readFileSync } from "node:fs";
import { join } from "node:path";
import { compareToFixture } from "../rules/compareToFixture";
import { extractDocModel, type Block, type DocModel, type ParagraphModel, type TableModel } from "../rules/docModel";
import { expectMatchesGolden } from "./__golden__/golden";
import { assembleDocumentBody, BODY_MARKER, type AssembleInput } from "./documentAssembler";
import { subjectLine } from "./titleBlock";

const fixture = JSON.parse(readFileSync(join(__dirname, "..", "..", "fixtures", "congvan-mau.json"), "utf8")) as DocModel;

/** Nội dung nhập giống file mẫu công văn. */
const sample: AssembleInput = {
  header: {
    parentLines: ["Sở Nông nghiệp và Môi trường", "Thành phố Cần Thơ"],
    issuer: "Chi cục Thủy lợi",
    sizes: { parent: 13, issuer: 13, quocHieu: 13, tieuNgu: 14, so: 13, placeDate: 13, subject: 13 },
    agencyLineRatio: 1 / 2,
    textWidthMm: 160,
    numberRow: { so: "Số:        /CCTL-ĐĐPCTT", placeDate: "Cần Thơ, ngày      tháng 10 năm 2026" },
    subjectRow: { text: subjectLine("cử công chức tham dự Hội nghị tập huấn") },
  },
  title: {
    variant: "letter",
    subject: "cử công chức tham dự Hội nghị tập huấn",
    recipients: ["Sở Nông nghiệp và Môi trường thành phố Cần Thơ"],
    sizes: { tenLoai: 14, trichYeu: 14, kinhGui: 14 },
    titleLineRatio: 1 / 2,
    textWidthMm: 160,
    lineMultiple: 1.17,
  },
  footer: {
    recipients: ["Như trên", "Chi cục trưởng (để b/c)", "Các đồng chí Mãnh, Thế Anh (để t/h)"],
    luu: "VT, CCTLTTA",
    authority: "KT.",
    positionLines: ["Chi cục trưởng", "Phó Chi cục trưởng"],
    signerName: "Mai Hiếu Hiền",
    signGapLines: 6,
    sizes: { noiNhanTieuDe: 12, noiNhanDanhSach: 11, quyenHanChucVu: 14, hoTen: 14 },
    textWidthMm: 160,
  },
};

function modelOf(input: AssembleInput): DocModel {
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${assembleDocumentBody(input)}</w:body></w:document>`;
  return extractDocModel({ documentXml: xml });
}

/** Các đoạn có chữ trong một bảng, theo thứ tự hàng, ô, đoạn (bỏ đoạn trống và đường kẻ). */
function textParagraphs(table: TableModel): ParagraphModel[] {
  return table.rows.flatMap((row) => row.flatMap((cell) => cell.filter((p) => p.text.trim() !== "")));
}

const asDoc = (blocks: Block[]): DocModel => ({ page: fixture.page, blocks });

describe("assembleDocumentBody: thứ tự các khối", () => {
  const blocks = modelOf(sample).blocks;

  it("khối đầu, Kính gửi, chỗ nội dung, khối cuối, đoạn đệm", () => {
    expect(blocks.map((b) => b.kind)).toEqual(["table", "paragraph", "paragraph", "table", "paragraph"]);
    expect((blocks[1] as ParagraphModel).text).toContain("Kính gửi:");
    expect((blocks[2] as ParagraphModel).text).toBe(BODY_MARKER);
  });

  it("chỗ đánh dấu nội dung xuất hiện đúng một lần", () => {
    expect(assembleDocumentBody(sample).split(BODY_MARKER)).toHaveLength(2);
  });

  it("không có phần tên loại thì chèn một đoạn trống để hai bảng không dính nhau", () => {
    const kinds = modelOf({ ...sample, title: undefined, header: { ...sample.header, subjectRow: undefined } }).blocks.map((b) => b.kind);
    expect(kinds).toEqual(["table", "paragraph", "paragraph", "table", "paragraph"]);
  });

  it("biến thể có tên loại", () => {
    const named: AssembleInput = {
      ...sample,
      header: { ...sample.header, subjectRow: undefined },
      title: { ...sample.title!, variant: "named", typeName: "Quyết định", subject: "Về việc ban hành Quy chế", recipients: [] },
    };
    const kinds = modelOf(named).blocks.map((b) => b.kind);
    expect(kinds).toEqual(["table", "paragraph", "paragraph", "paragraph", "paragraph", "table", "paragraph"]);
  });
});

describe("công văn nhập đúng nội dung của file mẫu khớp fixture về định dạng", () => {
  const blocks = modelOf(sample).blocks;
  const fixtureHead = fixture.blocks[0] as TableModel;
  const fixtureSign = fixture.blocks[8] as TableModel;

  it("trang: lề, khoảng cách header, footer (theo tuỳ chọn mặc định) nằm ở bước thiết lập trang, không ở thân", () => {
    expect(fixture.page.marginsMm).toEqual({ top: 20, bottom: 20, left: 30, right: 20 });
  });

  it("khối đầu: từng đoạn có chữ theo thứ tự khớp fixture (bỏ độ rộng cột và đường kẻ nổi của mẫu)", () => {
    const actual = textParagraphs(blocks[0] as TableModel);
    const expected = textParagraphs(fixtureHead);
    expect(actual).toHaveLength(expected.length);
    expect(compareToFixture(asDoc(actual), asDoc(expected), { ignore: [/floatingLinesMm/, /\.text$/] })).toEqual([]);
    expect(actual.map((p) => p.text)).toContain("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
  });

  it("Kính gửi khớp fixture", () => {
    const expected = fixture.blocks[2] as ParagraphModel;
    expect(compareToFixture(asDoc([blocks[1]]), asDoc([expected]), { ignore: [/\.text$/] })).toEqual([]);
  });

  it("khối cuối: nơi nhận và chữ ký khớp fixture (mẫu căn dàn đều ở dòng Nơi nhận)", () => {
    const actual = textParagraphs(blocks[3] as TableModel);
    const expected = textParagraphs(fixtureSign);
    expect(actual).toHaveLength(expected.length);
    expect(compareToFixture(asDoc(actual), asDoc(expected), { ignore: [/\.text$/, /\.align$/] })).toEqual([]);
  });
});

describe("golden toàn tài liệu", () => {
  it("công văn mẫu", () => {
    expectMatchesGolden("document-letter.xml", assembleDocumentBody(sample).replace(/></g, ">\n<"));
  });
});
