import { readFileSync } from "node:fs";
import { join } from "node:path";
import { buildRegistry, defaultValues } from "../settings/formatOptions";
import { ND30 } from "../standards/nd30/profile";
import type { DocModel, ParagraphModel, TableModel } from "./docModel";

/** fixtures/congvan-mau.json: đáp án chuẩn trích từ file mẫu công văn (T2A), đã ẩn danh. */
const raw = readFileSync(join(__dirname, "..", "..", "fixtures", "congvan-mau.json"), "utf8");
const fixture = JSON.parse(raw) as DocModel;
const defaults = defaultValues(buildRegistry(ND30));

const table = (index: number) => fixture.blocks[index] as TableModel;
const paragraph = (index: number) => fixture.blocks[index] as ParagraphModel;
const cell = (t: TableModel, row: number, col: number, index = 0) => t.rows[row][col][index];

describe("fixture công văn mẫu: không lộ dữ liệu cá nhân", () => {
  it("không có số điện thoại, email, họ tên hay tên cơ quan cụ thể", () => {
    expect(raw).not.toMatch(/\b0\d{9,10}\b/);
    expect(raw).not.toMatch(/@|\.vn\b|\.gov/i);
    expect(raw).not.toMatch(/Cần Thơ|Sở Nông|Chi cục|CHI CỤC|THÀNH PHỐ|Thông báo số|Kế hoạch số|Nguyễn|Trương/);
  });
});

describe("mặc định chuẩn Chính quyền khớp fixture công văn mẫu", () => {
  it("trang: lề, khoảng cách header, footer, trang đầu khác biệt, số trang ở header", () => {
    expect(fixture.page).toMatchObject({
      marginsMm: { top: defaults.marginTop, bottom: defaults.marginBottom, left: defaults.marginLeft, right: defaults.marginRight },
      headerDistanceMm: defaults.headerDistance,
      footerDistanceMm: defaults.footerDistance,
      titlePg: true,
      pageNumberInHeader: true,
    });
  });

  it("nội dung: cỡ, căn đều, thụt dòng đầu, giãn đoạn, giãn dòng", () => {
    expect(paragraph(3)).toMatchObject({
      sizePt: defaults.fontSize,
      align: "both",
      firstLineIndentMm: defaults.firstLineIndent,
      spaceBeforePt: defaults.spacing,
      spaceAfterPt: defaults.spacing,
      lineMultiple: defaults.lineSpacing,
      fontName: ND30.font.name,
    });
  });

  it("Kính gửi: căn giữa, giãn 12 pt (theo kế hoạch style của chuẩn)", () => {
    const plan = ND30.styles.elements.find((e) => e.elementId === "kinhGui");
    expect(paragraph(2)).toMatchObject({ align: "center", spaceBeforePt: plan?.spacingPt?.before, spaceAfterPt: plan?.spacingPt?.after, sizePt: defaults.fontSize });
  });

  it("bảng dữ liệu: cỡ chữ, giãn ô, hàng tiêu đề đậm căn giữa", () => {
    const data = table(5);
    expect(cell(data, 1, 1)).toMatchObject({ sizePt: defaults.tableFontSize, spaceBeforePt: defaults.tableSpacing, spaceAfterPt: defaults.tableSpacing });
    expect(cell(data, 0, 1)).toMatchObject({ bold: true, align: "center", sizePt: defaults.tableFontSize });
    expect(data.bordersNone).toBe(false);
  });

  it("bảng đầu văn bản: không viền, 2 cột khoảng 45% và 55%, 3 hàng", () => {
    const head = table(0);
    expect(head.bordersNone).toBe(true);
    expect(head.rows).toHaveLength(3);
    const [left, right] = head.columnsMm;
    expect(left / (left + right)).toBeCloseTo(0.45, 1);
  });

  it("cỡ chữ khối đầu văn bản: quốc hiệu, tiêu ngữ, cơ quan, số, địa danh, V/v", () => {
    const head = table(0);
    expect(cell(head, 0, 1, 0)).toMatchObject({ text: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", sizePt: defaults["size.quocHieu"], bold: true });
    expect(cell(head, 0, 1, 1)).toMatchObject({ text: "Độc lập - Tự do - Hạnh phúc", sizePt: defaults["size.tieuNgu"], bold: true });
    expect(cell(head, 0, 0, 2)).toMatchObject({ sizePt: defaults["size.tenCoQuanBanHanh"], bold: true });
    expect(cell(head, 1, 0)).toMatchObject({ sizePt: defaults["size.so"], align: "center", spaceBeforePt: 6 });
    expect(cell(head, 1, 1)).toMatchObject({ sizePt: defaults["size.diaDanhNgayThang"], italic: true });
    expect(cell(head, 2, 0)).toMatchObject({ sizePt: defaults["size.trichYeuCongVan"], align: "center" });
  });

  it("nơi nhận và chữ ký: cỡ chữ", () => {
    const sign = table(8);
    expect(cell(sign, 0, 0, 0)).toMatchObject({ text: "Nơi nhận:", sizePt: defaults["size.noiNhanTieuDe"], bold: true, italic: true });
    expect(cell(sign, 0, 1, 0)).toMatchObject({ sizePt: defaults["size.quyenHanChucVu"], bold: true });
  });
});
