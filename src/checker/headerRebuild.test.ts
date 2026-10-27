import { buildRegistry, defaultValues } from "../settings/formatOptions";
import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { buildRebuiltHeader, normalizeDate, normalizeSo, planHeaderRebuild, validateHeaderFields, type HeaderFields, type HeaderRebuildPlan } from "./headerRebuild";
import type { CheckDocument, CheckParagraph } from "./types";

interface P {
  text: string;
  t?: number;
  row?: number;
  col?: number;
  align?: "center" | "left" | "both" | "right";
  bold?: boolean;
}

function doc(items: P[]): CheckDocument {
  const paragraphs: CheckParagraph[] = items.map((p, index) => ({
    index,
    inTable: p.t !== undefined,
    table: p.t === undefined ? undefined : { index: p.t, row: p.row ?? 0, col: p.col ?? 0 },
    resolved: { kind: "paragraph", text: p.text, align: p.align, bold: p.bold },
  }));
  return { pages: [], paragraphs };
}

function plan(items: P[], profile = ND30): HeaderRebuildPlan {
  const result = planHeaderRebuild(doc(items), profile);
  if ("error" in result) throw new Error(result.error);
  return result;
}

describe("chuẩn hoá số, ký hiệu và ngày tháng", () => {
  it.each([
    ["Số 5 / PTC - KH", "Số: 05/PTC-KH"],
    ["Số: 12/BC-UBND", "Số: 12/BC-UBND"],
    ["Số:    /BC-UBND", "Số:      /BC-UBND"],
  ])("Chính quyền: %s → %s", (raw, expected) => {
    expect(normalizeSo(raw, ND30)).toBe(expected);
  });

  it.each([
    ["Số: 3-BC/ĐU", "Số 03-BC/ĐU"],
    ["Số 12 - HD / VPTW", "Số 12-HD/VPTW"],
  ])("Đảng: %s → %s", (raw, expected) => {
    expect(normalizeSo(raw, HD05)).toBe(expected);
  });

  it.each([
    ["Hà Nội ngày 5 tháng 1 năm 2026", "Hà Nội, ngày 05 tháng 01 năm 2026"],
    ["Cần Thơ,  ngày 15 tháng 03 năm 2026.", "Cần Thơ, ngày 15 tháng 3 năm 2026"],
    ["Phong Điền, ngày 9 tháng 12 năm 2026", "Phong Điền, ngày 09 tháng 12 năm 2026"],
  ])("%s → %s", (raw, expected) => {
    expect(normalizeDate(raw)).toBe(expected);
  });
});

describe("đọc khối đầu cũ chưa đúng quy định", () => {
  const congVanSai: P[] = [
    { text: "UBND HUYỆN PHONG ĐIỀN", t: 0, col: 0, align: "center" },
    { text: "PHÒNG TÀI CHÍNH - KẾ HOẠCH", t: 0, col: 0, bold: true, align: "center" },
    { text: "Số 5 / PTC - KH", t: 0, col: 0, align: "center" },
    { text: "V/v báo cáo kết quả", t: 0, col: 0, align: "center" },
    { text: "thu ngân sách quý I", t: 0, col: 0, align: "center" },
    { text: "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM", t: 0, col: 1, bold: true, align: "center" },
    { text: "Độc lập – Tự do – Hạnh phúc", t: 0, col: 1, bold: true, align: "center" },
    { text: "Phong Điền ngày 5 tháng 1 năm 2026", t: 0, col: 1, align: "center" },
    { text: "" },
    { text: "Kính gửi: Ủy ban nhân dân huyện Phong Điền", align: "center" },
    { text: "Thực hiện chỉ đạo của Ủy ban nhân dân huyện, Phòng báo cáo như sau:", align: "both" },
  ];

  it("công văn: tách cơ quan theo in đậm, chuẩn hoá số và ngày, gộp trích yếu nhiều dòng, lấy Kính gửi", () => {
    const p = plan(congVanSai);
    expect(p.fields).toEqual({
      parentLines: ["UBND HUYỆN PHONG ĐIỀN"],
      issuerLines: ["PHÒNG TÀI CHÍNH - KẾ HOẠCH"],
      so: "Số: 05/PTC-KH",
      placeDate: "Phong Điền, ngày 05 tháng 01 năm 2026",
      kind: "letter",
      typeName: "",
      subject: "Báo cáo kết quả thu ngân sách quý I",
      recipients: ["Ủy ban nhân dân huyện Phong Điền"],
    });
    expect(p.end).toBe(9);
    expect(p.tables).toEqual([0]);
  });

  it("văn bản có tên loại: tên loại theo bảng tên loại, trích yếu hai dòng, dòng kẻ dưới; thân giữ nguyên", () => {
    const p = plan([
      { text: "ỦY BAN NHÂN DÂN", t: 0, col: 0, bold: true },
      { text: "XÃ AN BÌNH", t: 0, col: 0, bold: true },
      { text: "Số: 7/BC-UBND", t: 0, row: 1, col: 0 },
      { text: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", t: 0, col: 1, bold: true },
      { text: "Độc lập - Tự do - Hạnh phúc", t: 0, col: 1, bold: true },
      { text: "An Bình, ngày 20 tháng 10 năm 2026", t: 0, row: 1, col: 1 },
      { text: "" },
      { text: "BÁO CÁO", align: "center", bold: true },
      { text: "Kết quả thực hiện nhiệm vụ", align: "center", bold: true },
      { text: "9 tháng đầu năm 2026.", align: "center", bold: true },
      { text: "_______________", align: "center" },
      { text: "I. KẾT QUẢ THỰC HIỆN", bold: true, align: "left" },
    ]);
    expect(p.fields).toMatchObject({
      parentLines: [],
      issuerLines: ["ỦY BAN NHÂN DÂN XÃ AN BÌNH"],
      so: "Số: 07/BC-UBND",
      kind: "named",
      typeName: "Báo cáo",
      subject: "Kết quả thực hiện nhiệm vụ 9 tháng đầu năm 2026",
      recipients: [],
    });
    expect(p.end).toBe(10);
  });

  it("khối đầu gõ bằng đoạn thường (không có bảng)", () => {
    const p = plan([
      { text: "SỞ NỘI VỤ", bold: true, align: "left" },
      { text: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", bold: true, align: "center" },
      { text: "Độc lập - Tự do - Hạnh phúc", align: "center" },
      { text: "Số: 15/TTr-SNV", align: "left" },
      { text: "Cần Thơ, ngày 3 tháng 2 năm 2026", align: "right" },
      { text: "TỜ TRÌNH", align: "center", bold: true },
      { text: "Về việc phê duyệt kế hoạch tuyển dụng", align: "center", bold: true },
      { text: "Kính gửi: Ủy ban nhân dân thành phố", align: "center" },
      { text: "Căn cứ Luật Cán bộ, công chức;", align: "both" },
    ]);
    expect(p.fields).toMatchObject({ issuerLines: ["SỞ NỘI VỤ"], so: "Số: 15/TTr-SNV", placeDate: "Cần Thơ, ngày 03 tháng 02 năm 2026", kind: "named", typeName: "Tờ trình", subject: "Về việc phê duyệt kế hoạch tuyển dụng", recipients: ["Ủy ban nhân dân thành phố"] });
    expect(p.tables).toEqual([]);
    expect(p.end).toBe(7);
  });

  it("nhiều nơi nhận gạch đầu dòng", () => {
    const p = plan([...congVanSai.slice(0, 9), { text: "Kính gửi:", align: "center" }, { text: "- Sở Tài chính;", align: "left" }, { text: "- Kho bạc Nhà nước.", align: "left" }, { text: "Thực hiện chỉ đạo…", align: "both" }]);
    expect(p.fields.recipients).toEqual(["Sở Tài chính", "Kho bạc Nhà nước"]);
    expect(p.end).toBe(11);
  });

  it("không có khối đầu thì báo, không đoán", () => {
    expect(planHeaderRebuild(doc([{ text: "Thực hiện chỉ đạo của cấp trên, đơn vị báo cáo như sau:", align: "both" }]), ND30)).toHaveProperty("error");
  });
});

describe("dựng khối đầu mới", () => {
  const values = defaultValues(buildRegistry(ND30));
  const fields: HeaderFields = {
    parentLines: ["UBND huyện Phong Điền"],
    issuerLines: ["Phòng Tài chính - Kế hoạch"],
    so: "Số: 05/PTC-KH",
    placeDate: "Phong Điền, ngày 05 tháng 01 năm 2026",
    kind: "named",
    typeName: "Báo cáo",
    subject: "Kết quả thu ngân sách quý I",
    recipients: [],
  };

  it("Chính quyền, có tên loại: quốc hiệu, tên cơ quan in hoa, tên loại in hoa, trích yếu", () => {
    const xml = buildRebuiltHeader(fields, ND30, values);
    for (const piece of ["CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc", "UBND HUYỆN PHONG ĐIỀN", "PHÒNG TÀI CHÍNH - KẾ HOẠCH", "Số: 05/PTC-KH", "BÁO CÁO", "Kết quả thu ngân sách quý I"]) expect(xml).toContain(piece);
  });

  it("Chính quyền, công văn: dòng V/v và Kính gửi trong bảng khối đầu", () => {
    const xml = buildRebuiltHeader({ ...fields, kind: "letter", typeName: "", recipients: ["Sở Tài chính"] }, ND30, values);
    expect(xml).toContain("V/v Kết quả thu ngân sách quý I");
    expect(xml).toContain("Kính gửi:");
    expect(xml).not.toContain("BÁO CÁO");
  });

  it("Chính quyền, tờ trình có Kính gửi: dòng Kính gửi dưới trích yếu", () => {
    expect(buildRebuiltHeader({ ...fields, typeName: "Tờ trình", recipients: ["Ủy ban nhân dân huyện"] }, ND30, values)).toContain("Kính gửi: Ủy ban nhân dân huyện");
  });

  it("Đảng, công văn: tiêu đề Đảng, dòng về việc", () => {
    const xml = buildRebuiltHeader({ ...fields, issuerLines: ["Đảng ủy xã An Bình"], parentLines: [], so: "Số 05-CV/ĐU", kind: "letter", typeName: "", recipients: ["Ban Tổ chức Huyện ủy"] }, HD05, defaultValues(buildRegistry(HD05)));
    expect(xml).toContain("ĐẢNG CỘNG SẢN VIỆT NAM");
    expect(xml).toContain("ĐẢNG ỦY XÃ AN BÌNH");
    expect(xml).toContain("về việc Kết quả thu ngân sách quý I");
  });

  it("thiếu cơ quan ban hành hoặc trích yếu thì không dựng", () => {
    expect(validateHeaderFields({ ...fields, issuerLines: [" "], subject: "" })).toEqual(["Hãy nhập tên cơ quan ban hành.", "Hãy nhập trích yếu nội dung."]);
    expect(() => buildRebuiltHeader({ ...fields, typeName: "" }, ND30, values)).toThrow("tên loại");
  });
});
