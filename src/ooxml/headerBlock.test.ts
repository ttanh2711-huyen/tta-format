import { expectMatchesGolden } from "./__golden__/golden";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { compareToFixture } from "../rules/compareToFixture";
import { extractDocModel, type DocModel, type ParagraphModel, type TableModel } from "../rules/docModel";
import { buildRegistry, defaultValues } from "../settings/formatOptions";
import { ND30 } from "../standards/nd30/profile";
import { buildHeaderBlock, headerInputFromOptions, validateHeaderInput, type HeaderBlockInput } from "./headerBlock";
import { textWidthMm } from "./textWidth";

const fixture = JSON.parse(readFileSync(join(__dirname, "..", "..", "fixtures", "congvan-mau.json"), "utf8")) as DocModel;
const fixtureHead = fixture.blocks[0] as TableModel;

const base: HeaderBlockInput = {
  parentLines: ["Sở Nông nghiệp và Môi trường", "Thành phố Cần Thơ"],
  issuer: "Chi cục Thủy lợi",
  sizes: { parent: 13, issuer: 13, quocHieu: 13, tieuNgu: 14, so: 13, placeDate: 13 },
  agencyLineRatio: 1 / 2,
  textWidthMm: 160,
};

function modelOf(input: HeaderBlockInput): DocModel {
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${buildHeaderBlock(input)}</w:body></w:document>`;
  return extractDocModel({ documentXml: xml });
}

const table = (input: HeaderBlockInput) => modelOf(input).blocks[0] as TableModel;
const asDoc = (p: ParagraphModel): DocModel => ({ page: fixture.page, blocks: [p] });

describe("buildHeaderBlock: cấu trúc", () => {
  it("một bảng 2 cột không viền, vừa vùng chữ, tỉ lệ 45% / 55%", () => {
    const t = table(base);
    expect(t.bordersNone).toBe(true);
    expect(t.rows).toHaveLength(1);
    expect(t.rows[0]).toHaveLength(2);
    expect(t.columnsMm[0] + t.columnsMm[1]).toBeCloseTo(160, 0);
    expect(t.columnsMm[0] / 160).toBeCloseTo(0.45, 2);
    expect(t.widthMm).toBeCloseTo(160, 0);
    expect(t.indentMm ?? 0).toBeCloseTo(0, 1);
  });

  it("ô trái: chủ quản (in hoa), cơ quan ban hành (in hoa), đường kẻ; ô phải: quốc hiệu, tiêu ngữ, đường kẻ", () => {
    const [left, right] = table(base).rows[0];
    expect(left.map((p) => p.text)).toEqual(["SỞ NÔNG NGHIỆP VÀ MÔI TRƯỜNG", "THÀNH PHỐ CẦN THƠ", "CHI CỤC THỦY LỢI", ""]);
    expect(right.map((p) => p.text)).toEqual(["CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc", ""]);
    expect(left[3].borderBottom).toBe(true);
    expect(right[2].borderBottom).toBe(true);
  });

  it.each([0, 1, 2, 3])("%i dòng cơ quan chủ quản đều sinh khối đúng", (count) => {
    const parents = ["Dòng một", "Dòng hai", "Dòng ba"].slice(0, count);
    const [left] = table({ ...base, parentLines: parents }).rows[0];
    expect(left).toHaveLength(count + 2);
    expect(left[count].text).toBe("CHI CỤC THỦY LỢI");
    expect(left[count].bold).toBe(true);
    for (const p of left.slice(0, count)) expect(p.bold).toBeFalsy();
  });

  it("bỏ dòng chủ quản trống và khoảng trắng thừa", () => {
    const [left] = table({ ...base, parentLines: ["  ", "Thành phố   Cần Thơ ", ""] }).rows[0];
    expect(left.map((p) => p.text)).toEqual(["THÀNH PHỐ CẦN THƠ", "CHI CỤC THỦY LỢI", ""]);
  });

  it("thoát ký tự đặc biệt của XML trong tên cơ quan", () => {
    const [left] = table({ ...base, parentLines: [], issuer: "Công ty A & B <test>" }).rows[0];
    expect(left[0].text).toBe("CÔNG TY A & B <TEST>");
  });
});

describe("buildHeaderBlock: đường kẻ", () => {
  function lineLengthMm(input: HeaderBlockInput, side: 0 | 1): number {
    const xml = buildHeaderBlock(input);
    const cells = xml.split("<w:tc>").slice(1);
    const ind = /<w:ind w:left="(\d+)" w:right="(\d+)"\/>/.exec(cells[side].slice(cells[side].lastIndexOf("<w:pBdr>")));
    if (!ind) throw new Error("không thấy thụt của đường kẻ");
    const widthTwips = Number(/<w:tcW w:w="(\d+)"/.exec(cells[side])?.[1]);
    return ((widthTwips - Number(ind[1]) - Number(ind[2])) * 25.4) / 1440;
  }

  it("đường kẻ dưới tên cơ quan dài đúng tỉ lệ so với dòng chữ", () => {
    const width = textWidthMm("CHI CỤC THỦY LỢI", 13, true);
    expect(lineLengthMm(base, 0)).toBeCloseTo(width * 0.5, 0);
    expect(lineLengthMm({ ...base, agencyLineRatio: 1 / 3 }, 0)).toBeCloseTo(width / 3, 0);
  });

  it("đường kẻ dưới tiêu ngữ dài bằng dòng tiêu ngữ", () => {
    expect(lineLengthMm(base, 1)).toBeCloseTo(textWidthMm("Độc lập - Tự do - Hạnh phúc", 14, true), 0);
  });

  it("tên cơ quan dài hơn ô thì đường kẻ tính theo bề rộng ô, không âm", () => {
    const long = "Sở Nông nghiệp và Môi trường thành phố Cần Thơ về việc thực hiện nhiệm vụ rất dài";
    const length = lineLengthMm({ ...base, parentLines: [], issuer: long }, 0);
    expect(length).toBeGreaterThan(0);
    expect(length).toBeLessThanOrEqual(72);
  });

  it("đường kẻ nằm trong đoạn riêng, tên cơ quan không bị thụt nên tự xuống dòng được", () => {
    const [left] = table({ ...base, issuer: "Chi cục Thủy lợi và Phòng chống thiên tai thành phố Cần Thơ" }).rows[0];
    const issuer = left[left.length - 2];
    expect(issuer.leftIndentMm).toBeUndefined();
    expect(issuer.borderBottom).toBeUndefined();
  });
});

describe("buildHeaderBlock: khớp fixture công văn mẫu", () => {
  const [left, right] = table(base).rows[0];
  const [fixLeft, fixRight] = fixtureHead.rows[0];
  const ignore = [/floatingLinesMm/, /\.text$/];
  const check = (actual: ParagraphModel, expected: ParagraphModel) => compareToFixture(asDoc(actual), asDoc(expected), { ignore });

  it("cơ quan chủ quản: căn giữa, 13 pt, không đậm", () => {
    expect(check(left[0], fixLeft[0])).toEqual([]);
    expect(check(left[1], fixLeft[1])).toEqual([]);
    expect(left[0].bold).toBeFalsy();
  });

  it("cơ quan ban hành: căn giữa, 13 pt, đậm", () => {
    expect(check(left[2], fixLeft[2])).toEqual([]);
  });

  it("quốc hiệu 13 pt đậm và tiêu ngữ 14 pt đậm, căn giữa", () => {
    expect(check(right[0], fixRight[0])).toEqual([]);
    expect(check(right[1], fixRight[1])).toEqual([]);
    expect(right[0].text).toBe(fixRight[0].text);
    expect(right[1].text).toBe(fixRight[1].text);
  });
});

describe("validateHeaderInput", () => {
  it("hợp lệ thì không có lỗi", () => {
    expect(validateHeaderInput(base)).toEqual([]);
  });

  it("thiếu cơ quan ban hành", () => {
    expect(validateHeaderInput({ ...base, issuer: "   " })).toEqual(["Hãy nhập tên cơ quan ban hành."]);
  });

  it("quá 3 dòng cơ quan chủ quản", () => {
    expect(validateHeaderInput({ ...base, parentLines: ["a", "b", "c", "d"] })).toEqual(["Cơ quan chủ quản tối đa 3 dòng."]);
  });
});

describe("buildHeaderBlock: golden", () => {
  it("khối đầu mẫu (2 dòng chủ quản) khớp file đối chiếu", () => {
    expectMatchesGolden("header-block.xml", buildHeaderBlock(base).replace(/></g, ">\n<"));
  });
});

describe("headerInputFromOptions", () => {
  const values = defaultValues(buildRegistry(ND30));

  it("lấy cỡ chữ, tỉ lệ đường kẻ và bề rộng vùng chữ từ Tuỳ chọn nhanh mặc định", () => {
    expect(headerInputFromOptions(ND30, values, { parentLines: ["A"], issuer: "B" })).toEqual({
      parentLines: ["A"],
      issuer: "B",
      sizes: { parent: 13, issuer: 13, quocHieu: 13, tieuNgu: 14, so: 13, placeDate: 13, subject: 13 },
      agencyLineRatio: 1 / 2,
      textWidthMm: 160,
    });
  });

  it("đổi lề thì vùng chữ và bảng đổi theo", () => {
    const wide = headerInputFromOptions(ND30, { ...values, marginRight: 15 }, { parentLines: [], issuer: "B" });
    expect(wide.textWidthMm).toBe(165);
  });
});

describe("buildHeaderBlock: hàng số và địa danh (T8)", () => {
  const withRow: HeaderBlockInput = { ...base, numberRow: { so: "Số:        /CCTL-ĐĐPCTT", placeDate: "Cần Thơ, ngày      tháng 10 năm 2026" } };

  it("thêm hàng thứ hai: số bên trái, địa danh và ngày bên phải", () => {
    const t = table(withRow);
    expect(t.rows).toHaveLength(2);
    expect(t.rows[1][0].map((p) => p.text)).toEqual(["Số:        /CCTL-ĐĐPCTT"]);
    expect(t.rows[1][1].map((p) => p.text)).toEqual(["Cần Thơ, ngày      tháng 10 năm 2026"]);
  });

  it("không truyền numberRow thì chỉ có một hàng", () => {
    expect(table(base).rows).toHaveLength(1);
  });

  it("khớp fixture: số căn giữa 13 pt giãn trước 6 pt; địa danh nghiêng 13 pt", () => {
    const t = table(withRow);
    expect(compareToFixture(asDoc(t.rows[1][0][0]), asDoc(fixtureHead.rows[1][0][0]), { ignore: [/\.text$/] })).toEqual([]);
    expect(compareToFixture(asDoc(t.rows[1][1][0]), asDoc(fixtureHead.rows[1][1][0]), { ignore: [/\.text$/] })).toEqual([]);
  });

  it("giữ nguyên khoảng trắng trong chỗ trống chờ điền", () => {
    expect(buildHeaderBlock(withRow)).toContain('xml:space="preserve">Số:        /CCTL-ĐĐPCTT<');
  });
});
