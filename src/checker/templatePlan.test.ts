import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { planTemplate } from "./templatePlan";
import type { CheckDocument, CheckParagraph } from "./types";

interface P {
  text: string;
  t?: number;
  row?: number;
  col?: number;
  align?: "center" | "left" | "both" | "right";
  bold?: boolean;
  italic?: boolean;
}

function doc(items: P[]): CheckDocument {
  const paragraphs: CheckParagraph[] = items.map((p, index) => ({
    index,
    inTable: p.t !== undefined,
    table: p.t === undefined ? undefined : { index: p.t, row: p.row ?? 0, col: p.col ?? 0 },
    resolved: { kind: "paragraph", text: p.text, align: p.align, bold: p.bold, italic: p.italic },
  }));
  return { pages: [], paragraphs };
}

const congVan: P[] = [
  { text: "CHI CỤC THỦY LỢI", t: 0, col: 0, bold: true, align: "center" },
  { text: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", t: 0, col: 1, bold: true, align: "center" },
  { text: "Độc lập - Tự do - Hạnh phúc", t: 0, col: 1, bold: true, align: "center" },
  { text: "Số: 05/CCTL-VP", t: 0, row: 1, col: 0, align: "center" },
  { text: "Cần Thơ, ngày 05 tháng 01 năm 2026", t: 0, row: 1, col: 1, italic: true, align: "center" },
  { text: "V/v cử công chức dự tập huấn", t: 0, row: 2, col: 0, align: "center" },
  { text: "Kính gửi:", t: 0, row: 3, col: 0, align: "right" },
  { text: "Sở Nông nghiệp và Môi trường", t: 0, row: 3, col: 1, align: "left" },
  { text: "Thực hiện Kế hoạch số 12/KH-SNN, Chi cục cử công chức tham dự.", align: "both" },
  { text: "Đề nghị Sở xem xét, chấp thuận.", align: "both" },
  { text: "Nơi nhận:", t: 1, col: 0, bold: true, italic: true },
  { text: "- Như trên;", t: 1, col: 0 },
  { text: "- Lưu: VT.", t: 1, col: 0 },
  { text: "KT. CHI CỤC TRƯỞNG", t: 1, col: 1, bold: true, align: "center" },
  { text: "PHÓ CHI CỤC TRƯỞNG", t: 1, col: 1, bold: true, align: "center" },
  { text: "", t: 1, col: 1 },
  { text: "Mai Hiếu Hiền", t: 1, col: 1, bold: true, align: "center" },
];

const tags = (items: P[], profile = ND30) => planTemplate(doc(items), profile).fields.map((f) => [f.tag, f.start, f.end]);

describe("lập kế hoạch biến văn bản thành mẫu (issue #2)", () => {
  it("công văn Chính quyền: số, ngày tháng, trích yếu, Kính gửi, thân, nơi nhận, quyền hạn chức vụ, họ tên", () => {
    expect(tags(congVan)).toEqual([
      ["so", 3, 3],
      ["diaDanhNgay", 4, 4],
      ["trichYeu", 5, 5],
      ["kinhGui", 7, 7],
      ["noiDung", 8, 9],
      ["noiNhan", 11, 12],
      ["chucVu", 13, 14],
      ["hoTen", 16, 16],
    ]);
  });

  it("chữ thay thế giữ đúng tiền tố thể thức; tên loại giữ nguyên chữ", () => {
    const fields = planTemplate(doc(congVan), ND30).fields;
    const by = (tag: string) => fields.find((f) => f.tag === tag)!;
    expect(by("so").placeholder).toBe("Số: …/…");
    expect(by("diaDanhNgay").placeholder).toBe("…, ngày … tháng … năm …");
    expect(by("trichYeu").placeholder).toBe("V/v …");
    expect(by("hoTen").placeholder).toBe("Họ và tên");
    expect(fields.every((f) => f.title !== "")).toBe(true);
  });

  it("văn bản có tên loại: tên loại giữ chữ, trích yếu, các căn cứ gộp một vùng, thân sau dòng chốt", () => {
    const qd: P[] = [
      ...congVan.slice(0, 5),
      { text: "QUYẾT ĐỊNH", align: "center", bold: true },
      { text: "Về việc ban hành quy chế làm việc", align: "center", bold: true },
      { text: "CHI CỤC TRƯỞNG CHI CỤC THỦY LỢI", align: "center", bold: true },
      { text: "Căn cứ Luật Tổ chức chính quyền địa phương;", italic: true, align: "both" },
      { text: "Xét đề nghị của Chánh Văn phòng,", italic: true, align: "both" },
      { text: "QUYẾT ĐỊNH:", align: "center", bold: true },
      { text: "Điều 1. Ban hành kèm theo Quyết định này Quy chế làm việc.", align: "both" },
      { text: "Điều 2. Quyết định có hiệu lực từ ngày ký.", align: "both" },
      ...congVan.slice(10),
    ];
    const plan = planTemplate(doc(qd), ND30);
    expect(plan.fields.map((f) => [f.tag, f.start, f.end, f.placeholder])).toEqual([
      ["so", 3, 3, "Số: …/…"],
      ["diaDanhNgay", 4, 4, "…, ngày … tháng … năm …"],
      ["tenLoai", 5, 5, null],
      ["trichYeu", 6, 6, "Về việc …"],
      ["canCu", 8, 9, "Căn cứ …;"],
      ["noiDung", 11, 12, "Nội dung văn bản"],
      ["noiNhan", 14, 15, "- …;"],
      ["chucVu", 16, 17, "QUYỀN HẠN, CHỨC VỤ NGƯỜI KÝ"],
      ["hoTen", 19, 19, "Họ và tên"],
    ]);
  });

  it("văn bản Đảng: số dạng \"Số …\", bỏ qua dòng 5 gạch nối", () => {
    const party: P[] = [
      { text: "TỈNH ỦY ĐỒNG THÁP", t: 0, col: 0, align: "center" },
      { text: "BAN TỔ CHỨC", t: 0, col: 0, bold: true, align: "center" },
      { text: "*", t: 0, col: 0, align: "center" },
      { text: "ĐẢNG CỘNG SẢN VIỆT NAM", t: 0, col: 1, bold: true, align: "center" },
      { text: "Số 05-BC/BTCTU", t: 0, row: 1, col: 0, align: "center" },
      { text: "Đồng Tháp, ngày 03 tháng 02 năm 2026", t: 0, row: 1, col: 1, italic: true, align: "right" },
      { text: "BÁO CÁO", align: "center", bold: true },
      { text: "kết quả hội nghị", align: "center", bold: true },
      { text: "-----", align: "center" },
      { text: "Hội nghị đã thống nhất các nội dung sau.", align: "both" },
      ...congVan.slice(10),
    ];
    const plan = planTemplate(doc(party), HD05);
    expect(plan.fields.find((f) => f.tag === "so")?.placeholder).toBe("Số …-…/…");
    expect(plan.fields.find((f) => f.tag === "noiDung")).toMatchObject({ start: 9, end: 9 });
  });

  it("thân có bảng thì không gộp thân (tránh làm hỏng bảng), báo lại", () => {
    const withTable: P[] = [...congVan.slice(0, 9), { text: "Ô bảng", t: 5 }, ...congVan.slice(9)];
    const plan = planTemplate(doc(withTable), ND30);
    expect(plan.fields.some((f) => f.tag === "noiDung")).toBe(false);
    expect(plan.notes.join(" ")).toContain("bảng");
  });

  it("tài liệu không có khối đầu thì không có gì để làm", () => {
    expect(planTemplate(doc([{ text: "Một đoạn thường." }]), ND30).fields).toEqual([]);
  });
});
