import { readFileSync } from "node:fs";
import { join } from "node:path";
import { compareToFixture } from "../rules/compareToFixture";
import { extractDocModel, type DocModel, type ParagraphModel, type TableModel } from "../rules/docModel";
import { buildRegistry, defaultValues } from "../settings/formatOptions";
import { ND30 } from "../standards/nd30/profile";
import { expectMatchesGolden } from "./__golden__/golden";
import { buildFooterBlock, findAcademicTitle, footerInputFromOptions, recipientLines, validateFooterInput, type FooterBlockInput } from "./footerBlock";

const fixture = JSON.parse(readFileSync(join(__dirname, "..", "..", "fixtures", "congvan-mau.json"), "utf8")) as DocModel;
const fixtureSign = fixture.blocks[8] as TableModel;

const base: FooterBlockInput = {
  recipients: ["Như trên", "Chi cục trưởng (để b/c)", "Các đồng chí Mãnh, Thế Anh (để t/h)"],
  luu: "VT, CCTLTTA",
  authority: "KT.",
  positionLines: ["Chi cục trưởng", "Phó Chi cục trưởng"],
  signerName: "Mai Hiếu Hiền",
  signGapLines: 6,
  sizes: { noiNhanTieuDe: 12, noiNhanDanhSach: 11, quyenHanChucVu: 14, hoTen: 14 },
  textWidthMm: 160,
};

function tableOf(input: FooterBlockInput): TableModel {
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${buildFooterBlock(input)}</w:body></w:document>`;
  return extractDocModel({ documentXml: xml }).blocks[0] as TableModel;
}

const asDoc = (p: ParagraphModel): DocModel => ({ page: fixture.page, blocks: [p] });

describe("recipientLines: danh sách nơi nhận", () => {
  it("gạch đầu dòng, ';' cuối mục, dấu chấm cuối mục cuối là Lưu", () => {
    expect(recipientLines(["Như trên", "Sở A"], "VT, ABC")).toEqual(["- Như trên;", "- Sở A;", "- Lưu: VT, ABC."]);
  });

  it("không có Lưu thì mục cuối của danh sách kết thúc bằng dấu chấm", () => {
    expect(recipientLines(["Như trên", "Sở A"], "")).toEqual(["- Như trên;", "- Sở A."]);
  });

  it("chuẩn hoá dấu gạch đầu dòng, khoảng trắng và dấu câu cuối người dùng đã gõ", () => {
    expect(recipientLines(["- Như trên;", "  –  Sở   A. ", ""], "  VT  ")).toEqual(["- Như trên;", "- Sở A;", "- Lưu: VT."]);
  });

  it("chữ 'Lưu:' người dùng đã gõ không bị lặp", () => {
    expect(recipientLines([], "Lưu: VT, ABC")).toEqual(["- Lưu: VT, ABC."]);
  });

  it("rỗng hết thì không có dòng nào", () => {
    expect(recipientLines(["", " "], " ")).toEqual([]);
  });
});

describe("findAcademicTitle: học hàm, học vị trong họ tên", () => {
  it.each(["TS. Nguyễn Văn A", "ThS. Trần B", "PGS.TS Lê C", "GS Phạm D", "BS CKI Hoàng E", "Th.S Võ F", "Tiến sĩ Đặng G", "thạc sĩ Bùi H", "KS. Đỗ I"])("%s có học hàm hoặc học vị", (name) => {
    expect(findAcademicTitle(name)).not.toBeNull();
  });

  it.each(["Mai Hiếu Hiền", "Trần Thị Bích", "Đoàn Văn Tiến", "Lê Thị Ths", "Nguyễn Ts"])("%s không có", (name) => {
    expect(findAcademicTitle(name)).toBeNull();
  });
});

describe("buildFooterBlock: cấu trúc và khớp fixture", () => {
  const table = tableOf(base);
  const [left, right] = table.rows[0];

  it("bảng 2 cột không viền vừa vùng chữ, ô trái hơi rộng hơn ô phải như mẫu", () => {
    expect(table.bordersNone).toBe(true);
    expect(table.rows).toHaveLength(1);
    expect(table.widthMm).toBeCloseTo(160, 0);
    expect(table.columnsMm[0]).toBeGreaterThan(table.columnsMm[1]);
    expect(table.columnsMm[0] / 160).toBeCloseTo(0.514, 2);
  });

  it("nơi nhận: tiêu đề 12 pt đậm nghiêng, danh sách 11 pt, mục Lưu ở cuối", () => {
    expect(left.map((p) => p.text)).toEqual(["Nơi nhận:", "- Như trên;", "- Chi cục trưởng (để b/c);", "- Các đồng chí Mãnh, Thế Anh (để t/h);", "- Lưu: VT, CCTLTTA."]);
    expect(compareToFixture(asDoc(left[0]), asDoc(fixtureSign.rows[0][0][0]), { ignore: [/align/] })).toEqual([]);
    for (const item of left.slice(1)) expect(item).toMatchObject({ sizePt: 11, fontName: "Times New Roman" });
    expect(left[1].bold).toBeFalsy();
  });

  it("chữ ký: quyền hạn và chức vụ in hoa đậm 14 pt căn giữa, 6 dòng trống, họ tên đậm", () => {
    expect(right.map((p) => p.text)).toEqual(["KT. CHI CỤC TRƯỞNG", "PHÓ CHI CỤC TRƯỞNG", "", "", "", "", "", "", "Mai Hiếu Hiền"]);
    expect(compareToFixture(asDoc(right[0]), asDoc(fixtureSign.rows[0][1][0]), { ignore: [/\.text$/] })).toEqual([]);
    expect(compareToFixture(asDoc(right[1]), asDoc(fixtureSign.rows[0][1][1]), { ignore: [/\.text$/] })).toEqual([]);
    expect(compareToFixture(asDoc(right[8]), asDoc(fixtureSign.rows[0][1][8]), { ignore: [/\.text$/] })).toEqual([]);
  });

  it("số dòng trống chỗ ký đổi được", () => {
    expect(tableOf({ ...base, signGapLines: 4 }).rows[0][1]).toHaveLength(2 + 4 + 1);
  });

  it("không có quyền hạn thì chức vụ in hoa một mình; chức vụ một dòng", () => {
    const [, r] = tableOf({ ...base, authority: "", positionLines: ["Giám đốc"] }).rows[0];
    expect(r[0].text).toBe("GIÁM ĐỐC");
    expect(r).toHaveLength(1 + 6 + 1);
  });

  it.each(["TM.", "Q.", "KT.", "TL.", "TUQ."])("quyền hạn %s đứng trước dòng chức vụ đầu", (authority) => {
    const [, r] = tableOf({ ...base, authority, positionLines: ["Chủ tịch"] }).rows[0];
    expect(r[0].text).toBe(`${authority} CHỦ TỊCH`);
  });

  it("không có nơi nhận thì ô trái chỉ có một đoạn trống", () => {
    const [l] = tableOf({ ...base, recipients: [], luu: "" }).rows[0];
    expect(l.map((p) => p.text)).toEqual([""]);
  });

  it("thoát ký tự XML trong họ tên và nơi nhận", () => {
    const t = tableOf({ ...base, recipients: ["A & B <c>"], signerName: "X & Y" });
    expect(t.rows[0][0][1].text).toBe("- A & B <c>;");
    expect(t.rows[0][1][t.rows[0][1].length - 1].text).toBe("X & Y");
  });

  it("golden", () => {
    expectMatchesGolden("footer-block.xml", buildFooterBlock(base).replace(/></g, ">\n<"));
  });
});

describe("validateFooterInput", () => {
  it("hợp lệ", () => {
    expect(validateFooterInput(base)).toEqual([]);
  });

  it("thiếu họ tên hoặc chức vụ", () => {
    expect(validateFooterInput({ ...base, signerName: " " })).toEqual(["Hãy nhập họ tên người ký."]);
    expect(validateFooterInput({ ...base, positionLines: ["", " "] })).toEqual(["Hãy nhập chức vụ người ký."]);
  });

  it("chức vụ tối đa hai dòng; số dòng trống từ 1 đến 12", () => {
    expect(validateFooterInput({ ...base, positionLines: ["a", "b", "c"] })).toEqual(["Chức vụ tối đa 2 dòng."]);
    expect(validateFooterInput({ ...base, signGapLines: 0 })).toEqual(["Số dòng trống chỗ ký phải từ 1 đến 12."]);
  });
});

describe("footerInputFromOptions", () => {
  it("lấy cỡ chữ và số dòng trống từ Tuỳ chọn nhanh", () => {
    const values = defaultValues(buildRegistry(ND30));
    expect(footerInputFromOptions(ND30, values, { recipients: [], luu: "VT", authority: "", positionLines: ["a"], signerName: "b" })).toMatchObject({
      sizes: { noiNhanTieuDe: 12, noiNhanDanhSach: 11, quyenHanChucVu: 14, hoTen: 14 },
      signGapLines: 6,
      textWidthMm: 160,
    });
  });
});
