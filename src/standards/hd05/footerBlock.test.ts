import { expectMatchesGolden } from "../../ooxml/__golden__/golden";
import { extractDocModel, type DocModel, type TableModel } from "../../rules/docModel";
import { buildRegistry, defaultValues } from "../../settings/formatOptions";
import { HD05 } from "./profile";
import { buildPartyFooterBlock, partyFooterFromOptions, validatePartyFooter, type PartyFooterInput } from "./footerBlock";

const base: PartyFooterInput = {
  recipients: ["Như trên", "Các đồng chí Ủy viên Ban Thường vụ"],
  luu: "VT, TCTU",
  drafter: "",
  authority: "K/T",
  positionLines: ["Trưởng ban", "Phó Trưởng ban"],
  signerName: "Nguyễn Văn A",
  confirmation: "",
  signGapLines: 6,
  sizes: { noiNhanTu: 14, noiNhanDanhSach: 12, quyenHan: 14, chucVu: 14, hoTen: 14, kyHieuNguoiSoan: 11 },
  textWidthMm: 165,
};

function xmlOf(input: PartyFooterInput): string {
  return buildPartyFooterBlock(input);
}

function tableOf(input: PartyFooterInput): TableModel {
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${xmlOf(input)}</w:body></w:document>`;
  return (extractDocModel({ documentXml: xml }) as DocModel).blocks[0] as TableModel;
}

/** Các run (theo thứ tự) của đoạn có chữ `text` trong XML: danh sách [chữ, có đậm]. */
function runsOf(xml: string, text: string): Array<[string, boolean]> {
  const paragraph = xml.match(/<w:p>[\s\S]*?<\/w:p>/g)?.find((p) => p.includes(text));
  if (!paragraph) throw new Error(`Không thấy đoạn chứa "${text}"`);
  return [...paragraph.matchAll(/<w:r>([\s\S]*?)<\/w:r>/g)].map((m) => [/<w:t[^>]*>([^<]*)<\/w:t>/.exec(m[1])?.[1] ?? "", m[1].includes("<w:b/>")]);
}

describe("buildPartyFooterBlock: nơi nhận", () => {
  const [left] = tableOf(base).rows[0];

  it("bảng 2 cột không viền vừa vùng chữ", () => {
    const t = tableOf(base);
    expect(t.bordersNone).toBe(true);
    expect(t.widthMm).toBeCloseTo(165, 0);
    expect(t.rows).toHaveLength(1);
    expect(t.rows[0]).toHaveLength(2);
  });

  it("'Nơi nhận:' 14 pt đứng, gạch chân; danh sách 12 pt, ';' cuối mục, mục Lưu kết thúc '.'", () => {
    expect(left.map((p) => p.text)).toEqual(["Nơi nhận:", "- Như trên;", "- Các đồng chí Ủy viên Ban Thường vụ;", "- Lưu: VT, TCTU."]);
    expect(left[0]).toMatchObject({ sizePt: 14, underline: true });
    expect(left[0].bold).toBeFalsy();
    expect(left[0].italic).toBeFalsy();
    for (const item of left.slice(1)) expect(item).toMatchObject({ sizePt: 12, align: "left" });
  });

  it("ký hiệu người soạn và số bản phát hành 11 pt ở cuối nơi nhận khi có nhập", () => {
    const [l] = tableOf({ ...base, drafter: "NVA (15 bản)" }).rows[0];
    expect(l[l.length - 1]).toMatchObject({ text: "NVA (15 bản)", sizePt: 11 });
  });

  it("không có nơi nhận thì ô trái chỉ có một đoạn trống", () => {
    expect(tableOf({ ...base, recipients: [], luu: "" }).rows[0][0].map((p) => p.text)).toEqual([""]);
  });
});

describe("buildPartyFooterBlock: chữ ký", () => {
  it("quyền hạn đậm cùng dòng với chức vụ in hoa không đậm; dòng chức vụ thứ hai không đậm", () => {
    const xml = xmlOf(base);
    expect(runsOf(xml, "TRƯỞNG BAN")[0]).toEqual(["K/T", true]);
    expect(runsOf(xml, "TRƯỞNG BAN").some(([text, bold]) => text.includes("TRƯỞNG BAN") && !bold)).toBe(true);
    expect(runsOf(xml, "PHÓ TRƯỞNG BAN")).toEqual([["PHÓ TRƯỞNG BAN", false]]);
  });

  it("chỗ ký 6 dòng trống rồi họ tên đậm in như đã nhập", () => {
    const [, right] = tableOf(base).rows[0];
    expect(right.map((p) => p.text)).toEqual(["K/T TRƯỞNG BAN", "PHÓ TRƯỞNG BAN", "", "", "", "", "", "", "Nguyễn Văn A"]);
    expect(right[right.length - 1]).toMatchObject({ bold: true, sizePt: 14, align: "center" });
    expect(right[1]).toMatchObject({ sizePt: 14, align: "center" });
    expect(right[1].bold).toBeFalsy();
  });

  it("số dòng trống đổi được", () => {
    expect(tableOf({ ...base, signGapLines: 4 }).rows[0][1]).toHaveLength(2 + 4 + 1);
  });

  it("người ký có hai chức danh: chức danh thứ hai in hoa dưới chỗ ký, trên họ tên", () => {
    const [, right] = tableOf({ ...base, authority: "", positionLines: ["Trưởng ban"], secondTitle: "Bí thư Tỉnh ủy", confirmation: "", signGapLines: 4 }).rows[0];
    expect(right.map((p) => p.text)).toEqual(["TRƯỞNG BAN", "", "", "", "", "BÍ THƯ TỈNH ỦY", base.signerName]);
  });

  it.each(["T/M", "K/T", "T/L", "Q."])("quyền hạn %s đứng trước dòng chức vụ đầu", (authority) => {
    const [, right] = tableOf({ ...base, authority, positionLines: ["Bí thư"] }).rows[0];
    expect(right[0].text).toBe(`${authority} BÍ THƯ`);
  });

  it("không quyền hạn thì chức vụ in hoa không đậm một mình", () => {
    const xml = xmlOf({ ...base, authority: "", positionLines: ["Bí thư"] });
    expect(runsOf(xml, "BÍ THƯ")).toEqual([["BÍ THƯ", false]]);
  });

  it("dòng xác nhận chữ ký bật tắt được và nằm sau họ tên", () => {
    const off = tableOf(base).rows[0][1].map((p) => p.text);
    expect(off.some((t) => t.includes("Xác nhận"))).toBe(false);
    const on = tableOf({ ...base, confirmation: "Trần Văn B, Chánh Văn phòng" }).rows[0][1].map((p) => p.text);
    expect(on[on.length - 1]).toBe("Xác nhận chữ ký của đồng chí Trần Văn B, Chánh Văn phòng");
    expect(on[on.length - 2]).toBe("Nguyễn Văn A");
  });

  it("thoát ký tự XML", () => {
    const t = tableOf({ ...base, recipients: ["A & B"], signerName: "X & Y" });
    expect(t.rows[0][0][1].text).toBe("- A & B;");
    expect(t.rows[0][1][t.rows[0][1].length - 1].text).toBe("X & Y");
  });

  it("golden", () => {
    expectMatchesGolden("party-footer-block.xml", xmlOf({ ...base, drafter: "NVA (15 bản)", confirmation: "Trần Văn B" }).replace(/></g, ">\n<"));
  });
});

describe("validatePartyFooter", () => {
  it("hợp lệ", () => {
    expect(validatePartyFooter(base)).toEqual([]);
  });

  it("thiếu họ tên hoặc chức vụ, quá hai dòng chức vụ, số dòng trống ngoài khoảng", () => {
    expect(validatePartyFooter({ ...base, signerName: " " })).toEqual(["Hãy nhập họ tên người ký."]);
    expect(validatePartyFooter({ ...base, positionLines: [" "] })).toEqual(["Hãy nhập chức vụ người ký."]);
    expect(validatePartyFooter({ ...base, positionLines: ["a", "b", "c"] })).toEqual(["Chức vụ tối đa 2 dòng."]);
    expect(validatePartyFooter({ ...base, signGapLines: 0 })).toEqual(["Số dòng trống chỗ ký phải từ 1 đến 12."]);
  });
});

describe("partyFooterFromOptions", () => {
  it("cỡ chữ và số dòng trống theo Tuỳ chọn nhanh mặc định của chuẩn Đảng", () => {
    const values = defaultValues(buildRegistry(HD05));
    expect(partyFooterFromOptions(HD05, values, { recipients: [], luu: "", drafter: "", authority: "", positionLines: ["a"], signerName: "b", confirmation: "" })).toMatchObject({
      sizes: { noiNhanTu: 14, noiNhanDanhSach: 12, quyenHan: 14, chucVu: 14, hoTen: 14, kyHieuNguoiSoan: 11 },
      signGapLines: 6,
      textWidthMm: 165,
    });
  });
});
