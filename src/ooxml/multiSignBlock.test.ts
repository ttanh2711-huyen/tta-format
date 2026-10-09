import { expectMatchesGolden } from "./__golden__/golden";
import { extractDocModel, type TableModel } from "../rules/docModel";
import { buildMultiSignBlock, MAX_PARTIES, parseSignLines, validateMultiSign, type MultiSignInput, type SignParty } from "./multiSignBlock";

const party = (title: string, name = "", position: string[] = []): SignParty => ({ title, positionLines: position, signerName: name });
const base = (parties: SignParty[], extra: Partial<MultiSignInput> = {}): MultiSignInput => ({
  parties,
  signGapLines: 4,
  sizes: { title: 14, name: 14, noiNhanTieuDe: 12, noiNhanDanhSach: 11 },
  textWidthMm: 160,
  recipients: [],
  luu: "",
  ...extra,
});

function table(xml: string): TableModel {
  const doc = extractDocModel({ documentXml: `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${xml}</w:body></w:document>` });
  return doc.blocks[0] as TableModel;
}

const A = party("Đại diện bên A", "Nguyễn Văn A", ["Giám đốc"]);
const B = party("Đại diện bên B", "Trần Thị B", ["Tổng giám đốc"]);
const C = party("Đại diện bên C", "Lê Văn C");
const D = party("Đại diện bên D", "Phạm Thị D");

describe("kiểm tra đầu vào", () => {
  it("2 đến 4 bên hợp lệ", () => {
    for (const parties of [[A, B], [A, B, C], [A, B, C, D]]) expect(validateMultiSign(base(parties))).toEqual([]);
  });

  it("1 bên hoặc 5 bên bị từ chối", () => {
    expect(validateMultiSign(base([A]))).not.toEqual([]);
    expect(validateMultiSign(base([A, B, C, D, A]))).not.toEqual([]);
    expect(() => buildMultiSignBlock(base([A]))).toThrow();
    expect(MAX_PARTIES).toBe(4);
  });

  it("bên thiếu tên, chức vụ quá 3 dòng, số dòng ký ngoài khoảng", () => {
    expect(validateMultiSign(base([party(" "), B]))).not.toEqual([]);
    expect(validateMultiSign(base([party("A", "", ["1", "2", "3", "4"]), B]))).not.toEqual([]);
    expect(validateMultiSign(base([A, B], { signGapLines: 0 }))).not.toEqual([]);
    expect(validateMultiSign(base([A, B], { signGapLines: 13 }))).not.toEqual([]);
  });
});

describe("bố cục", () => {
  it.each([2, 3, 4])("%i bên: bảng một hàng, các cột bằng nhau, tổng bằng vùng chữ", (count) => {
    const t = table(buildMultiSignBlock(base([A, B, C, D].slice(0, count))));
    expect(t.rows).toHaveLength(1);
    expect(t.rows[0]).toHaveLength(count);
    expect(t.bordersNone).toBe(true);
    const widths = t.columnsMm;
    expect(widths).toHaveLength(count);
    expect(Math.max(...widths) - Math.min(...widths)).toBeLessThanOrEqual(0.2);
    expect(widths.reduce((a, b) => a + b, 0)).toBeCloseTo(160, 0);
  });

  it("mỗi cột: tên bên và chức vụ in hoa đậm, chỗ ký, họ tên đậm", () => {
    const t = table(buildMultiSignBlock(base([A, B])));
    const column = t.rows[0][0].map((p) => ({ text: p.text, bold: p.bold, align: p.align }));
    expect(column.slice(0, 2)).toEqual([{ text: "ĐẠI DIỆN BÊN A", bold: true, align: "center" }, { text: "GIÁM ĐỐC", bold: true, align: "center" }]);
    expect(column[column.length - 1]).toEqual({ text: "Nguyễn Văn A", bold: true, align: "center" });
    expect(column.filter((p) => p.text === "")).toHaveLength(4);
  });

  it("tên bên rất dài vẫn nằm trong cột của nó (không thêm cột, không thêm hàng)", () => {
    const long = party("Đại diện Công ty trách nhiệm hữu hạn một thành viên đầu tư xây dựng và phát triển hạ tầng", "Nguyễn Văn Rất Dài Họ Tên", ["Phó Tổng giám đốc phụ trách kinh doanh"]);
    const t = table(buildMultiSignBlock(base([long, B, C, D])));
    expect(t.rows).toHaveLength(1);
    expect(t.rows[0]).toHaveLength(4);
  });

  it("họ tên để trống thì chừa chỗ, không lỗi", () => {
    expect(table(buildMultiSignBlock(base([party("Chủ tọa"), party("Thư ký")]))).rows[0]).toHaveLength(2);
  });

  it("có nơi nhận thì thêm hàng trải hết các cột; không có thì không có hàng nơi nhận", () => {
    const withRecipients = buildMultiSignBlock(base([A, B], { recipients: ["Như trên", "Các bên"], luu: "VT" }));
    expect(withRecipients).toContain('<w:gridSpan w:val="2"/>');
    expect(table(withRecipients).rows).toHaveLength(2);
    expect(buildMultiSignBlock(base([A, B]))).not.toContain("gridSpan");
  });

  it.each([2, 3, 4])("golden %i bên", (count) => {
    expectMatchesGolden(`multiSign-${count}.xml`, buildMultiSignBlock(base([A, B, C, D].slice(0, count), count === 2 ? { recipients: ["Như trên"], luu: "VT" } : {})));
  });
});

describe("xác nhận chữ ký (chuẩn Đảng)", () => {
  it("có xác nhận thì thêm dòng dưới họ tên, không đậm; không có thì không thêm", () => {
    const withConfirm = table(buildMultiSignBlock(base([{ ...A, confirmation: "Trần Văn B, Bí thư" }, B])));
    const column = withConfirm.rows[0][0];
    expect(column[column.length - 1]).toMatchObject({ text: "Xác nhận chữ ký của đồng chí Trần Văn B, Bí thư", align: "center" });
    expect(column[column.length - 1].bold).not.toBe(true);
    expect(withConfirm.rows[0][1].some((p) => p.text.startsWith("Xác nhận"))).toBe(false);
  });

  it("đọc cột thứ tư của dòng nhập là xác nhận", () => {
    expect(parseSignLines("Người ghi | | Nguyễn A | Trần B, Bí thư")[0]).toEqual({ title: "Người ghi", positionLines: [], signerName: "Nguyễn A", confirmation: "Trần B, Bí thư" });
  });
});

describe("đọc các dòng nhập", () => {
  it("Tên bên | Chức vụ | Họ tên, bỏ dòng trống", () => {
    expect(parseSignLines("Đại diện bên A | Giám đốc | Nguyễn A\n\nChủ tọa | |\nThư ký")).toEqual([
      { title: "Đại diện bên A", positionLines: ["Giám đốc"], signerName: "Nguyễn A" },
      { title: "Chủ tọa", positionLines: [], signerName: "" },
      { title: "Thư ký", positionLines: [], signerName: "" },
    ]);
  });
});
