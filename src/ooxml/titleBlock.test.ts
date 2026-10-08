import { readFileSync } from "node:fs";
import { join } from "node:path";
import { compareToFixture } from "../rules/compareToFixture";
import { extractDocModel, type DocModel, type ParagraphModel, type TableModel } from "../rules/docModel";
import { expectMatchesGolden } from "./__golden__/golden";
import { buildHeaderBlock, type HeaderBlockInput } from "./headerBlock";
import { textWidthMm } from "./textWidth";
import { buildTitleBody, subjectLine, validateTitleInput, type TitleBlockInput } from "./titleBlock";

const fixture = JSON.parse(readFileSync(join(__dirname, "..", "..", "fixtures", "congvan-mau.json"), "utf8")) as DocModel;
const fixtureHead = fixture.blocks[0] as TableModel;
const fixtureKinhGui = fixture.blocks[2] as ParagraphModel;

const sizes = { tenLoai: 14, trichYeu: 14, kinhGui: 14 };

const named: TitleBlockInput = { variant: "named", typeName: "Quyết định", subject: "Về việc ban hành Quy chế làm việc", recipients: [], sizes, titleLineRatio: 1 / 2, textWidthMm: 160, lineMultiple: 1.17 };
const letter: TitleBlockInput = { variant: "letter", subject: "cử công chức tham dự Hội nghị tập huấn", recipients: ["Sở Nông nghiệp và Môi trường thành phố Cần Thơ"], sizes, titleLineRatio: 1 / 2, textWidthMm: 160, lineMultiple: 1.17 };

function paragraphs(input: TitleBlockInput): ParagraphModel[] {
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${buildTitleBody(input)}</w:body></w:document>`;
  return extractDocModel({ documentXml: xml }).blocks as ParagraphModel[];
}

describe("buildTitleBody: có tên loại", () => {
  it("tên loại in hoa đậm căn giữa, trích yếu đậm căn giữa, rồi đường kẻ", () => {
    const [type, subject, line] = paragraphs(named);
    expect(type).toMatchObject({ text: "QUYẾT ĐỊNH", bold: true, align: "center", sizePt: 14 });
    expect(subject).toMatchObject({ text: "Về việc ban hành Quy chế làm việc", bold: true, align: "center", sizePt: 14 });
    expect(line).toMatchObject({ text: "", borderBottom: true });
  });

  it("đường kẻ dưới trích yếu dài đúng tỉ lệ so với dòng chữ", () => {
    const width = textWidthMm("Về việc ban hành Quy chế làm việc", 14, true);
    for (const ratio of [1 / 3, 1 / 2]) {
      const xml = buildTitleBody({ ...named, titleLineRatio: ratio });
      const ind = /<w:ind w:left="(\d+)" w:right="(\d+)"\/>/.exec(xml);
      const lineMm = ((9071 - Number(ind?.[1]) - Number(ind?.[2])) * 25.4) / 1440;
      expect(lineMm).toBeCloseTo(width * ratio, 0);
    }
  });

  it("trích yếu nhiều dòng (dài hơn vùng chữ) vẫn là một đoạn tự xuống dòng, đường kẻ không vượt vùng chữ", () => {
    const long = `Về việc ${"thực hiện nhiệm vụ trọng tâm ".repeat(8)}`.trim();
    const blocks = paragraphs({ ...named, subject: long });
    expect(blocks).toHaveLength(3);
    expect(blocks[1].text).toBe(long);
    const xml = buildTitleBody({ ...named, subject: long });
    const ind = /<w:ind w:left="(\d+)" w:right="(\d+)"\/>/.exec(xml);
    expect(Number(ind?.[1])).toBeGreaterThanOrEqual(0);
  });

  it("gọn khoảng trắng thừa, thoát ký tự XML", () => {
    const blocks = paragraphs({ ...named, typeName: " nghị  quyết ", subject: "A & B <x>" });
    expect(blocks[0].text).toBe("NGHỊ QUYẾT");
    expect(blocks[1].text).toBe("A & B <x>");
  });

  it("golden: quyết định", () => {
    expectMatchesGolden("title-named.xml", buildTitleBody(named).replace(/></g, ">\n<"));
  });
});

describe("buildTitleBody: công văn (Kính gửi)", () => {
  it("một nơi nhận: một dòng căn giữa, giãn trước và sau 12 pt, khớp fixture", () => {
    const blocks = paragraphs(letter);
    expect(blocks).toHaveLength(1);
    expect(blocks[0].text).toBe("Kính gửi: Sở Nông nghiệp và Môi trường thành phố Cần Thơ");
    expect(compareToFixture({ page: fixture.page, blocks: [blocks[0]] }, { page: fixture.page, blocks: [fixtureKinhGui] }, { ignore: [/\.text$/] })).toEqual([]);
  });

  it("nhiều nơi nhận: dòng 'Kính gửi:' rồi mỗi nơi một dòng gạch đầu dòng, chấm phẩy, cuối cùng chấm", () => {
    const blocks = paragraphs({ ...letter, recipients: ["Bộ Nội vụ", "Bộ Tài chính", "UBND các tỉnh"] });
    expect(blocks.map((b) => b.text)).toEqual(["Kính gửi:", "- Bộ Nội vụ;", "- Bộ Tài chính;", "- UBND các tỉnh."]);
    expect(blocks[0].align).toBe("center");
    expect(blocks[1].align).toBe("left");
    expect(blocks[1].leftIndentMm).toBeGreaterThan(0);
  });

  it("bỏ dòng nơi nhận trống, bỏ dấu câu thừa người dùng gõ cuối dòng", () => {
    const blocks = paragraphs({ ...letter, recipients: ["", " Bộ Nội vụ; ", "Bộ Tài chính."] });
    expect(blocks.map((b) => b.text)).toEqual(["Kính gửi:", "- Bộ Nội vụ;", "- Bộ Tài chính."]);
  });

  it("golden: công văn một nơi nhận và nhiều nơi nhận", () => {
    expectMatchesGolden("title-letter.xml", buildTitleBody(letter).replace(/></g, ">\n<"));
    expectMatchesGolden("title-letter-multi.xml", buildTitleBody({ ...letter, recipients: ["Bộ Nội vụ", "Bộ Tài chính"] }).replace(/></g, ">\n<"));
  });
});

describe("subjectLine: V/v của công văn", () => {
  it("thêm tiền tố V/v", () => {
    expect(subjectLine("cử công chức  tham dự")).toBe("V/v cử công chức tham dự");
  });

  it("không thêm hai lần nếu người dùng đã gõ", () => {
    expect(subjectLine("V/v cử công chức")).toBe("V/v cử công chức");
    expect(subjectLine("v/v cử công chức")).toBe("V/v cử công chức");
  });
});

describe("hàng V/v trong bảng đầu văn bản", () => {
  const head: HeaderBlockInput = {
    parentLines: [],
    issuer: "Chi cục Thủy lợi",
    sizes: { parent: 13, issuer: 13, quocHieu: 13, tieuNgu: 14, so: 13, placeDate: 13, subject: 13 },
    agencyLineRatio: 1 / 2,
    textWidthMm: 160,
    numberRow: { so: "Số: 05/CCTL", placeDate: "Cần Thơ, ngày 05 tháng 10 năm 2026" },
    subjectRow: { text: subjectLine("cử công chức tham dự") },
  };

  it("hàng thứ ba: V/v ở ô trái, căn giữa 13 pt không đậm giãn trước 6 pt, khớp fixture; ô phải trống", () => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${buildHeaderBlock(head)}</w:body></w:document>`;
    const table = extractDocModel({ documentXml: xml }).blocks[0] as TableModel;
    expect(table.rows).toHaveLength(3);
    const subject = table.rows[2][0][0];
    expect(subject.text).toBe("V/v cử công chức tham dự");
    expect(subject.bold).toBeFalsy();
    expect(compareToFixture({ page: fixture.page, blocks: [subject] }, { page: fixture.page, blocks: [fixtureHead.rows[2][0][0]] }, { ignore: [/\.text$/] })).toEqual([]);
    expect(table.rows[2][1].map((p) => p.text)).toEqual([""]);
  });
});

describe("validateTitleInput", () => {
  it("hợp lệ", () => {
    expect(validateTitleInput(named)).toEqual([]);
    expect(validateTitleInput(letter)).toEqual([]);
  });

  it("có tên loại: cần tên loại và trích yếu", () => {
    expect(validateTitleInput({ ...named, typeName: " " })).toEqual(["Hãy nhập tên loại văn bản."]);
    expect(validateTitleInput({ ...named, subject: "" })).toEqual(["Hãy nhập trích yếu nội dung."]);
  });

  it("công văn: cần trích yếu V/v và ít nhất một nơi nhận", () => {
    expect(validateTitleInput({ ...letter, subject: "  " })).toEqual(["Hãy nhập trích yếu nội dung."]);
    expect(validateTitleInput({ ...letter, recipients: ["", "  "] })).toEqual(["Hãy nhập ít nhất một nơi nhận (Kính gửi)."]);
  });
});
