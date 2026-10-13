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
  it("một bảng 2 cột không viền, vừa vùng chữ, cột phải đủ rộng để quốc hiệu nằm trên một dòng", () => {
    const t = table(base);
    expect(t.bordersNone).toBe(true);
    expect(t.rows).toHaveLength(1);
    expect(t.rows[0]).toHaveLength(2);
    expect(t.columnsMm[0] + t.columnsMm[1]).toBeCloseTo(160, 0);
    expect(t.columnsMm[1]).toBeGreaterThan(textWidthMm("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", 13, true));
    expect(t.columnsMm[0] / 160).toBeGreaterThanOrEqual(0.3);
    expect(t.widthMm).toBeCloseTo(160, 0);
    expect(t.indentMm ?? 0).toBeCloseTo(0, 1);
  });

  it("ô trái: chủ quản, cơ quan ban hành (in hoa); ô phải: quốc hiệu, tiêu ngữ; đường kẻ là hình đường thẳng neo trong đoạn chữ", () => {
    const [left, right] = table(base).rows[0];
    expect(left.map((p) => p.text)).toEqual(["SỞ NÔNG NGHIỆP VÀ MÔI TRƯỜNG", "THÀNH PHỐ CẦN THƠ", "CHI CỤC THỦY LỢI"]);
    expect(right.map((p) => p.text)).toEqual(["CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc"]);
    expect(left[2].floatingLinesMm).toHaveLength(1);
    expect(right[1].floatingLinesMm).toHaveLength(1);
    expect(left[0].floatingLinesMm).toBeUndefined();
    expect(left[2].borderBottom).toBeUndefined();
  });

  it.each([0, 1, 2, 3])("%i dòng cơ quan chủ quản đều sinh khối đúng", (count) => {
    const parents = ["Dòng một", "Dòng hai", "Dòng ba"].slice(0, count);
    const [left] = table({ ...base, parentLines: parents }).rows[0];
    expect(left).toHaveLength(count + 1);
    expect(left[count].text).toBe("CHI CỤC THỦY LỢI");
    expect(left[count].bold).toBe(true);
    for (const p of left.slice(0, count)) expect(p.bold).toBeFalsy();
  });

  it("bỏ dòng chủ quản trống và khoảng trắng thừa", () => {
    const [left] = table({ ...base, parentLines: ["  ", "Thành phố   Cần Thơ ", ""] }).rows[0];
    expect(left.map((p) => p.text)).toEqual(["THÀNH PHỐ CẦN THƠ", "CHI CỤC THỦY LỢI"]);
  });

  it("thoát ký tự đặc biệt của XML trong tên cơ quan", () => {
    const [left] = table({ ...base, parentLines: [], issuer: "Công ty A & B <test>" }).rows[0];
    expect(left[0].text).toBe("CÔNG TY A & B <TEST>");
  });
});

describe("buildHeaderBlock: đường kẻ (Insert > Shapes > Line)", () => {
  const lineMm = (input: HeaderBlockInput, side: 0 | 1): number => {
    const cell = table(input).rows[0][side];
    return cell[cell.length - 1].floatingLinesMm?.[0] ?? Number.NaN;
  };

  it("đường kẻ dưới tên cơ quan dài đúng tỉ lệ so với dòng chữ", () => {
    const width = textWidthMm("CHI CỤC THỦY LỢI", 13, true);
    expect(lineMm(base, 0)).toBeCloseTo(width * 0.5, 0);
    expect(lineMm({ ...base, agencyLineRatio: 1 / 3 }, 0)).toBeCloseTo(width / 3, 0);
  });

  it("đường kẻ dưới tiêu ngữ dài bằng dòng tiêu ngữ", () => {
    expect(lineMm(base, 1)).toBeCloseTo(textWidthMm("Độc lập - Tự do - Hạnh phúc", 14, true), 0);
  });

  it("tên cơ quan dài hơn ô thì đường kẻ tính theo bề rộng ô", () => {
    const long = "Sở Nông nghiệp và Môi trường thành phố Cần Thơ về việc thực hiện nhiệm vụ rất dài";
    const length = lineMm({ ...base, parentLines: [], issuer: long }, 0);
    expect(length).toBeGreaterThan(0);
    expect(length).toBeLessThanOrEqual(72);
  });

  it("đường kẻ là hình nổi neo vào đoạn chữ, canh giữa cột, không dùng viền hay đoạn riêng", () => {
    const xml = buildHeaderBlock(base);
    expect(xml.match(/prst="line"/g)).toHaveLength(2);
    expect(xml).toContain('<wp:positionH relativeFrom="column">');
    expect(xml).toContain('<wp:positionV relativeFrom="paragraph">');
    expect(xml).not.toContain("<w:pBdr>");
    expect(new Set(xml.match(/<wp:docPr id="(\d+)"/g)).size).toBe(2);
  });

  it("tên cơ quan dài tự xuống dòng, không thụt", () => {
    const [left] = table({ ...base, issuer: "Chi cục Thủy lợi và Phòng chống thiên tai thành phố Cần Thơ" }).rows[0];
    expect(left[left.length - 1].leftIndentMm).toBeUndefined();
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

describe("buildHeaderBlock: hàng Kính gửi (nhãn canh phải ở ô trái, nơi nhận canh trái ở ô phải)", () => {
  const withRow: HeaderBlockInput = {
    ...base,
    numberRow: { so: "Số: 05/UBND-VP", placeDate: "Cần Thơ, ngày 05 tháng 10 năm 2026" },
    subjectRow: { text: "V/v cử công chức" },
    recipientsRow: { label: "Kính gửi:", lines: ["- Bộ Nội vụ;", "- Bộ Tài chính."], sizePt: 14 },
  };

  it("là hàng thứ tư của bảng, ngay dưới ô V/v", () => {
    const t = table(withRow);
    expect(t.rows).toHaveLength(4);
    expect(t.rows[2][0][0].text).toBe("V/v cử công chức");
  });

  it("cả hai ô bắt đầu từ dòng thứ hai (dòng đầu để trống); từ hai nơi nhận trở lên, danh sách nằm dưới dòng của nhãn một dòng", () => {
    const [left, right] = table(withRow).rows[3];
    expect(left.map((p) => p.text)).toEqual(["", "Kính gửi:"]);
    expect(right.map((p) => p.text)).toEqual(["", "", "- Bộ Nội vụ;", "- Bộ Tài chính."]);
  });

  it("một nơi nhận: nằm ngang hàng với nhãn Kính gửi", () => {
    const [left, right] = table({ ...withRow, recipientsRow: { label: "Kính gửi:", lines: ["Sở Nông nghiệp"], sizePt: 14 } }).rows[3];
    expect(left.map((p) => p.text)).toEqual(["", "Kính gửi:"]);
    expect(right.map((p) => p.text)).toEqual(["", "Sở Nông nghiệp"]);
  });

  it("nhãn canh phải 14 pt; nơi nhận canh trái 14 pt, xuống dòng thẳng hàng sau dấu gạch", () => {
    const [left, right] = table(withRow).rows[3];
    expect(left[1]).toMatchObject({ align: "right", sizePt: 14 });
    expect(right[2]).toMatchObject({ align: "left", sizePt: 14 });
    expect(right[2].leftIndentMm).toBeGreaterThan(0);
  });

  it("một nơi nhận không có gạch đầu dòng, không thụt treo", () => {
    const [, right] = table({ ...withRow, recipientsRow: { label: "Kính gửi:", lines: ["Sở Nông nghiệp"], sizePt: 14 } }).rows[3];
    expect(right[1].text).toBe("Sở Nông nghiệp");
    expect(right[1].leftIndentMm).toBeUndefined();
  });

  it("golden", () => {
    expectMatchesGolden("header-block-recipients.xml", buildHeaderBlock(withRow).replace(/></g, ">\n<"));
  });
});
