import { expectMatchesGolden } from "../../ooxml/__golden__/golden";
import { textWidthMm } from "../../ooxml/textWidth";
import { extractDocModel, type DocModel, type TableModel } from "../../rules/docModel";
import { buildRegistry, defaultValues } from "../../settings/formatOptions";
import { HD05 } from "./profile";
import { buildPartyHeaderBlock, partyHeaderFromOptions, validatePartyHeader, type PartyHeaderInput } from "./headerBlock";

const base: PartyHeaderInput = {
  parentLines: ["Tỉnh ủy Đồng Tháp"],
  issuerLines: ["Ban Tổ chức"],
  sizes: { title: 15, parent: 14, issuer: 14, so: 14, placeDate: 14 },
  textWidthMm: 165,
};

function tableOf(input: PartyHeaderInput): TableModel {
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${buildPartyHeaderBlock(input)}</w:body></w:document>`;
  return (extractDocModel({ documentXml: xml }) as DocModel).blocks[0] as TableModel;
}

describe("buildPartyHeaderBlock", () => {
  it("một bảng 2 cột không viền vừa vùng chữ", () => {
    const t = tableOf(base);
    expect(t.bordersNone).toBe(true);
    expect(t.rows).toHaveLength(1);
    expect(t.rows[0]).toHaveLength(2);
    expect(t.widthMm).toBeCloseTo(165, 0);
    expect(t.columnsMm[0] / 165).toBeCloseTo(0.45, 2);
  });

  it("trái: cơ quan cấp trên in hoa không đậm 14 pt, cơ quan ban hành in hoa đậm 14 pt, dấu * ngay dưới", () => {
    const [left] = tableOf(base).rows[0];
    expect(left.map((p) => p.text)).toEqual(["TỈNH ỦY ĐỒNG THÁP", "BAN TỔ CHỨC", "*"]);
    expect(left[0]).toMatchObject({ sizePt: 14, align: "center" });
    expect(left[0].bold).toBeFalsy();
    expect(left[1]).toMatchObject({ sizePt: 14, bold: true, align: "center" });
    expect(left[2]).toMatchObject({ align: "center", sizePt: 14 });
    expect(left[2].bold).toBeFalsy();
  });

  it("phải: tiêu đề in hoa đậm 15 pt, đường kẻ là hình đường thẳng neo trong đoạn, dài bằng tiêu đề", () => {
    const [, right] = tableOf(base).rows[0];
    expect(right).toHaveLength(1);
    expect(right[0]).toMatchObject({ text: "ĐẢNG CỘNG SẢN VIỆT NAM", sizePt: 15, bold: true, align: "center" });
    expect(right[0].floatingLinesMm?.[0]).toBeCloseTo(textWidthMm("ĐẢNG CỘNG SẢN VIỆT NAM", 15, true), 0);
    expect(right[0].borderBottom).toBeUndefined();
    expect(buildPartyHeaderBlock(base)).not.toContain("<w:pBdr>");
  });

  it.each([0, 1, 2])("%i dòng cơ quan cấp trên đều sinh khối đúng", (count) => {
    const parents = ["Tỉnh ủy", "Huyện ủy"].slice(0, count);
    const [left] = tableOf({ ...base, parentLines: parents }).rows[0];
    expect(left).toHaveLength(count + 2);
    expect(left[count].text).toBe("BAN TỔ CHỨC");
    expect(left[left.length - 1].text).toBe("*");
  });

  it("cơ quan ban hành nhiều dòng đều đậm, dấu * nằm sau dòng cuối", () => {
    const [left] = tableOf({ ...base, parentLines: [], issuerLines: ["Ủy ban Kiểm tra", "Ban Tổ chức"] }).rows[0];
    expect(left.map((p) => p.text)).toEqual(["ỦY BAN KIỂM TRA", "BAN TỔ CHỨC", "*"]);
    expect(left[0].bold).toBe(true);
    expect(left[1].bold).toBe(true);
  });

  it("liên cơ quan nối bằng dấu gạch nối trong một dòng được giữ nguyên", () => {
    const [left] = tableOf({ ...base, parentLines: [], issuerLines: ["Ban Tổ chức - Ban Tuyên giáo"] }).rows[0];
    expect(left[0].text).toBe("BAN TỔ CHỨC - BAN TUYÊN GIÁO");
  });

  it("bỏ dòng trống, gọn khoảng trắng, thoát ký tự XML", () => {
    const [left] = tableOf({ ...base, parentLines: ["", "  "], issuerLines: ["  Ban   A & B ", ""] }).rows[0];
    expect(left.map((p) => p.text)).toEqual(["BAN A & B", "*"]);
  });

  it("tên cơ quan dài xuống dòng trong ô, không thụt", () => {
    const [left] = tableOf({ ...base, issuerLines: ["Ban Chấp hành Đảng bộ Khối các cơ quan và doanh nghiệp tỉnh Đồng Tháp nhiệm kỳ 2025-2030"] }).rows[0];
    expect(left[1].leftIndentMm).toBeUndefined();
  });

  it("golden: hai dòng cấp trên, cơ quan ban hành một dòng", () => {
    expectMatchesGolden("party-header-block.xml", buildPartyHeaderBlock({ ...base, parentLines: ["Tỉnh ủy Đồng Tháp", "Huyện ủy Cao Lãnh"] }).replace(/></g, ">\n<"));
  });

  it("golden: liên cơ quan nhiều dòng, không có cấp trên", () => {
    expectMatchesGolden("party-header-block-multi.xml", buildPartyHeaderBlock({ ...base, parentLines: [], issuerLines: ["Ủy ban Kiểm tra", "Ban Tổ chức"] }).replace(/></g, ">\n<"));
  });
});

describe("validatePartyHeader", () => {
  it("hợp lệ", () => {
    expect(validatePartyHeader(base)).toEqual([]);
  });

  it("thiếu cơ quan ban hành", () => {
    expect(validatePartyHeader({ ...base, issuerLines: [" ", ""] })).toEqual(["Hãy nhập tên cơ quan ban hành."]);
  });

  it("quá số dòng", () => {
    expect(validatePartyHeader({ ...base, parentLines: ["a", "b", "c"] })).toEqual(["Cơ quan cấp trên tối đa 2 dòng."]);
    expect(validatePartyHeader({ ...base, issuerLines: ["a", "b", "c", "d"] })).toEqual(["Cơ quan ban hành tối đa 3 dòng."]);
  });
});

describe("partyHeaderFromOptions", () => {
  it("cỡ chữ theo Tuỳ chọn nhanh mặc định và vùng chữ theo lề chuẩn Đảng (165 mm)", () => {
    const values = defaultValues(buildRegistry(HD05));
    expect(partyHeaderFromOptions(HD05, values, { parentLines: [], issuerLines: ["a"] })).toEqual({
      parentLines: [],
      issuerLines: ["a"],
      sizes: { title: 15, parent: 14, issuer: 14, so: 14, placeDate: 14 },
      textWidthMm: 165,
    });
  });
});

describe("buildPartyHeaderBlock: hàng số, ký hiệu và ngày tháng (TD2)", () => {
  const withRow: PartyHeaderInput = { ...base, numberRow: { so: "Số 05-HD/VPTW", placeDate: "Hà Nội, ngày 03 tháng 02 năm 2026" } };

  it("thêm hàng thứ hai: số căn giữa 14 pt giãn trước 6 pt, ngày nghiêng 14 pt", () => {
    const t = tableOf(withRow);
    expect(t.rows).toHaveLength(2);
    expect(t.rows[1][0][0]).toMatchObject({ text: "Số 05-HD/VPTW", sizePt: 14, align: "center", spaceBeforePt: 6 });
    expect(t.rows[1][0][0].italic).toBeFalsy();
    expect(t.rows[1][1][0]).toMatchObject({ text: "Hà Nội, ngày 03 tháng 02 năm 2026", sizePt: 14, italic: true, align: "center" });
  });

  it("không truyền numberRow thì chỉ một hàng", () => {
    expect(tableOf(base).rows).toHaveLength(1);
  });

  it("giữ nguyên chỗ trống chờ điền", () => {
    expect(buildPartyHeaderBlock({ ...withRow, numberRow: { so: "Số       -HD/VPTW", placeDate: "Hà Nội, ngày      tháng      năm 2026" } })).toContain('xml:space="preserve">Số       -HD/VPTW<');
  });

  it("golden", () => {
    expectMatchesGolden("party-header-block-number.xml", buildPartyHeaderBlock(withRow).replace(/></g, ">\n<"));
  });
});

describe("buildPartyHeaderBlock: hàng \"về việc\" của công văn (TD3)", () => {
  it("hàng thứ ba: dòng về việc nghiêng 12 pt ở ô trái, ô phải trống", () => {
    const t = tableOf({ ...base, numberRow: { so: "Số 05-CV/VPTW", placeDate: "Hà Nội, ngày 03 tháng 02 năm 2026" }, subjectRow: { text: "về việc cử cán bộ", sizePt: 12 } });
    expect(t.rows).toHaveLength(3);
    expect(t.rows[2][0][0]).toMatchObject({ text: "về việc cử cán bộ", sizePt: 12, italic: true, align: "center" });
    expect(t.rows[2][1].map((p) => p.text)).toEqual([""]);
  });
});

describe("buildPartyHeaderBlock: hàng Kính gửi của công văn", () => {
  it("hàng cuối: nhãn nghiêng canh phải ở ô trái, nơi nhận canh trái ở ô phải, cả hai bắt đầu từ dòng thứ hai", () => {
    const t = tableOf({
      ...base,
      numberRow: { so: "Số 05-CV/VPTW", placeDate: "Hà Nội, ngày 03 tháng 02 năm 2026" },
      subjectRow: { text: "về việc cử cán bộ", sizePt: 12 },
      recipientsRow: { label: "Kính gửi:", italicLabel: true, lines: ["- Ban A;", "- Ban B."], sizePt: 14 },
    });
    expect(t.rows).toHaveLength(4);
    const [left, right] = t.rows[3];
    expect(left.map((p) => p.text)).toEqual(["", "Kính gửi:"]);
    expect(left[1]).toMatchObject({ align: "right", italic: true, sizePt: 14 });
    expect(right.map((p) => p.text)).toEqual(["", "", "- Ban A;", "- Ban B."]);
    expect(right[2]).toMatchObject({ align: "left", sizePt: 14 });
    expect(right[2].italic).toBeFalsy();
  });
});
