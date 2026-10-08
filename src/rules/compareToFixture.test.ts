import type { DocModel } from "./docModel";
import { compareToFixture } from "./compareToFixture";

const page = { widthMm: 210, heightMm: 297, marginsMm: { top: 20, bottom: 20, left: 30, right: 20 }, headerDistanceMm: 12.5, footerDistanceMm: 5.3, titlePg: true, pageNumberInHeader: true };

function model(): DocModel {
  return {
    page,
    blocks: [
      { kind: "paragraph", text: "Kính gửi: […]", align: "center", sizePt: 14, spaceBeforePt: 12, spaceAfterPt: 12 },
      {
        kind: "table",
        columnsMm: [82.9, 100.5],
        bordersNone: true,
        rows: [[[{ kind: "paragraph", text: "[…]", sizePt: 13, bold: true }], [{ kind: "paragraph", text: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", sizePt: 13, bold: true }]]],
      },
    ],
  };
}

describe("compareToFixture", () => {
  it("giống nhau thì không có khác biệt", () => {
    expect(compareToFixture(model(), model())).toEqual([]);
  });

  it("chỉ so các thuộc tính fixture có khai báo; thuộc tính thừa của tài liệu không bị coi là lỗi", () => {
    const actual = model();
    (actual.blocks[0] as { italic?: boolean }).italic = true;
    expect(compareToFixture(actual, model())).toEqual([]);
  });

  it("báo khác biệt kèm đường dẫn và giá trị hai bên", () => {
    const actual = model();
    (actual.blocks[0] as { sizePt: number }).sizePt = 13;
    expect(compareToFixture(actual, model())).toEqual(["blocks[0].sizePt: tài liệu 13, mẫu 14"]);
  });

  it("thuộc tính fixture có mà tài liệu thiếu thì báo thiếu", () => {
    const actual = model();
    delete (actual.blocks[0] as { align?: string }).align;
    expect(compareToFixture(actual, model())).toEqual(["blocks[0].align: tài liệu không có, mẫu \"center\""]);
  });

  it("khác số khối, số hàng, số ô", () => {
    const actual = model();
    actual.blocks.pop();
    expect(compareToFixture(actual, model())).toEqual(["blocks: tài liệu có 1 phần tử, mẫu có 2"]);
  });

  it("sai khác số nhỏ trong dung sai thì bỏ qua (làm tròn đổi đơn vị)", () => {
    const actual = model();
    actual.page = { ...page, headerDistanceMm: 12.52 };
    expect(compareToFixture(actual, model())).toEqual([]);
  });

  it("so chữ thể thức đúng từng ký tự, so nhãn khi mẫu đã ẩn danh", () => {
    const actual = model();
    const table = actual.blocks[1] as { rows: Array<Array<Array<{ text: string }>>> };
    table.rows[0][0][0].text = "SỞ NÔNG NGHIỆP";
    table.rows[0][1][0].text = "CỘNG HÒA";
    (actual.blocks[0] as { text: string }).text = "Kính gửi: Sở Nông nghiệp";
    expect(compareToFixture(actual, model())).toEqual([
      'blocks[1].rows[0][1][0].text: tài liệu "CỘNG HÒA", mẫu "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"',
    ]);
  });

  it("nhãn đầu dòng của mẫu ẩn danh vẫn phải khớp", () => {
    const actual = model();
    (actual.blocks[0] as { text: string }).text = "Nơi gửi: Sở";
    expect(compareToFixture(actual, model())).toEqual(['blocks[0].text: tài liệu "Nơi gửi: Sở", mẫu bắt đầu bằng "Kính gửi:"']);
  });

  it("bỏ qua các đường dẫn được chỉ định", () => {
    const actual = model();
    (actual.blocks[1] as { columnsMm: number[] }).columnsMm = [70, 100];
    expect(compareToFixture(actual, model()).length).toBeGreaterThan(0);
    expect(compareToFixture(actual, model(), { ignore: [/^blocks\[1\]\.columnsMm/] })).toEqual([]);
  });
});
