import { buildCheckDocument } from "../../../checker/documentModel";
import { runChecks } from "../../../checker/engine";
import { headerFooterRules } from "../../../checker/ruleSets";
import { expectMatchesGolden } from "../../../ooxml/__golden__/golden";
import { BODY_MARKER } from "../../../ooxml/documentAssembler";
import { multiSignFromOptions } from "../../../ooxml/multiSignBlock";
import { parseStyleTable } from "../../../ooxml/styleResolver";
import { buildStyleSpecs } from "../../../ooxml/styleSpecs";
import { toStyleDefinitionsXml } from "../../../ooxml/styleXml";
import { extractDocModel, type ParagraphModel, type TableModel } from "../../../rules/docModel";
import { templatesFor } from "../../../templates/registry";
import { SMALL_FORMS, findSmallForm } from "../../../templates/smallForms";
import type { PartyAssembleInput } from "../assembleDocument";
import { HD05 } from "../profile";
import { assemblePartyCongressBody, assemblePartyFormBody, assemblePartyJointBody, assemblePartyMinutesBody, buildCongressName } from "./documents";

const sizes = { theLoai: 16, trichYeu: 14, trichYeuCongVan: 12, kinhGui: 14 };
const input = (typeName: string, soAbbr = "BB"): PartyAssembleInput => ({
  header: { parentLines: ["Tỉnh ủy Đồng Tháp"], issuerLines: ["Ban Tổ chức"], sizes: { title: 15, parent: 14, issuer: 14, so: 14, placeDate: 14 }, numberRow: { so: `Số 05-${soAbbr}/BTCTU`, placeDate: "Đồng Tháp, ngày 03 tháng 02 năm 2026" }, textWidthMm: 165 },
  title: { variant: "coTenLoai", typeName, subject: "kết quả hội nghị", recipients: [], basis: [], sizes, lineExactPt: 20, textWidthMm: 165 },
  footer: {
    recipients: ["Như trên"],
    luu: "VT",
    drafter: "",
    authority: "T/M",
    positionLines: ["Đoàn chủ tịch"],
    signerName: "Nguyễn Văn A",
    confirmation: "Trần Văn B, Bí thư",
    signGapLines: 6,
    sizes: { noiNhanTu: 14, noiNhanDanhSach: 12, quyenHan: 14, chucVu: 14, hoTen: 14, kyHieuNguoiSoan: 11 },
    textWidthMm: 165,
  },
});

function issuesOf(bodyXml: string) {
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml.replace(BODY_MARKER, "")}</w:body></w:document>`;
  const styles = parseStyleTable(`<w:styles>${toStyleDefinitionsXml(buildStyleSpecs(HD05))}</w:styles>`);
  const page = { widthMm: 210, heightMm: 297, marginsMm: { ...HD05.pageDefaults.marginsMm }, headerDistanceMm: HD05.pageDefaults.headerDistanceMm, footerDistanceMm: HD05.pageDefaults.footerDistanceMm, mirrorMargins: false };
  return runChecks(buildCheckDocument(extractDocModel({ documentXml: xml }), styles, [page]), HD05, headerFooterRules(HD05));
}

const sign = (parties: Parameters<typeof multiSignFromOptions>[2]["parties"], recipients: string[] = []) => multiSignFromOptions(HD05, {}, { parties, recipients, luu: "" });
const model = (xml: string) => extractDocModel({ documentXml: `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${xml}</w:body></w:document>` });

describe("biên bản của Đảng", () => {
  const parties = [
    { title: "Người ghi biên bản", positionLines: [], signerName: "Lê Thị C", confirmation: "Trần Văn B, Bí thư" },
    { title: "Chủ trì hội nghị", positionLines: ["Bí thư"], signerName: "Nguyễn Văn A" },
  ];

  it("hai chữ ký, dòng xác nhận chữ ký khi không đóng dấu, các mục đầu biên bản; 0 lỗi", () => {
    const xml = assemblePartyMinutesBody(input("Biên bản"), "HD05_", sign(parties));
    const doc = model(xml);
    const signTable = doc.blocks[doc.blocks.length - 2] as TableModel;
    expect(signTable.rows[0]).toHaveLength(2);
    expect(signTable.rows[0][0].some((p) => p.text === "Xác nhận chữ ký của đồng chí Trần Văn B, Bí thư")).toBe(true);
    const texts = doc.blocks.filter((b): b is ParagraphModel => b.kind === "paragraph").map((p) => p.text);
    expect(texts).toEqual(expect.arrayContaining(["BIÊN BẢN", `Thời gian: ${BODY_MARKER}`, "Nội dung:"]));
    expect(issuesOf(xml)).toEqual([]);
  });

  it("có Nơi nhận kiểu Đảng (gạch chân, không đậm) vẫn 0 lỗi", () => {
    const xml = assemblePartyMinutesBody(input("Biên bản"), "HD05_", sign(parties, ["Như trên", "Lưu hồ sơ"]));
    expect(xml).toContain('<w:u w:val="single"/>');
    expect(issuesOf(xml)).toEqual([]);
  });

  it("công văn bị từ chối", () => {
    expect(() => assemblePartyMinutesBody({ ...input("Biên bản"), title: { ...input("x").title!, variant: "congVan" } }, "HD05_", sign(parties))).toThrow();
  });
});

describe("văn bản liên cơ quan", () => {
  const two = [
    { title: "T/M Ban Thường vụ Tỉnh ủy", positionLines: ["Bí thư"], signerName: "Nguyễn Văn A" },
    { title: "T/M Ủy ban Kiểm tra Tỉnh ủy", positionLines: ["Chủ nhiệm"], signerName: "Lê Văn C" },
  ];

  it("hai khối ký cạnh nhau, tên cơ quan nối bằng gạch nối ở ký hiệu; 0 lỗi", () => {
    const base = input("Quy chế", "QC");
    const joint = { ...base, header: { ...base.header, issuerLines: ["Ban Tổ chức Tỉnh ủy", "Ủy ban Kiểm tra Tỉnh ủy"], numberRow: { so: "Số 05-QC/BTCTU-UBKT", placeDate: "Đồng Tháp, ngày 03 tháng 02 năm 2026" } } };
    const xml = assemblePartyJointBody(joint, sign(two));
    const signTable = model(xml).blocks.filter((b): b is TableModel => b.kind === "table").pop()!;
    expect(signTable.rows[0]).toHaveLength(2);
    expect(issuesOf(xml)).toEqual([]);
  });

  it("đúng hai khối ký", () => {
    expect(() => assemblePartyJointBody(input("Quy chế"), sign([two[0], two[1], two[0]]))).toThrow();
    expect(() => assemblePartyJointBody(input("Quy chế"), { ...sign(two), parties: [two[0]] })).toThrow();
  });
});

describe("văn bản đại hội", () => {
  it("tên đại hội in hoa đậm ở trên thể loại; ký T/M Đoàn chủ tịch có xác nhận chữ ký; 0 lỗi", () => {
    const xml = assemblePartyCongressBody(input("Nghị quyết", "NQ"), "Đại hội đại biểu Đảng bộ tỉnh lần thứ XI", 14);
    const paragraphs = model(xml).blocks.filter((b): b is ParagraphModel => b.kind === "paragraph");
    expect(paragraphs[0]).toMatchObject({ text: "ĐẠI HỘI ĐẠI BIỂU ĐẢNG BỘ TỈNH LẦN THỨ XI", bold: true, align: "center" });
    expect(paragraphs[1].text).toBe("NGHỊ QUYẾT");
    expect(issuesOf(xml)).toEqual([]);
  });

  it("thiếu tên đại hội bị từ chối", () => {
    expect(() => buildCongressName("  ", 14)).toThrow();
    expect(() => assemblePartyCongressBody(input("Nghị quyết"), "", 14)).toThrow();
  });
});

describe("giấy tờ hành chính nhỏ của chuẩn Đảng", () => {
  const forms = templatesFor(HD05).filter((entry) => entry.family === "form");

  it("có đủ tám giấy tờ", () => {
    expect(forms.map((entry) => entry.docTypeId).sort()).toEqual(["giay-chung-nhan", "giay-di-duong", "giay-gioi-thieu", "giay-moi", "giay-nghi-phep", "phieu-chuyen", "phieu-gui", "thu-cong"]);
  });

  it.each(forms.map((entry) => [entry.docTypeId, entry.typeName] as const))("%s: 0 lỗi", (id, typeName) => {
    const form = findSmallForm(id)!;
    expect(SMALL_FORMS).toContain(form);
    expect(issuesOf(assemblePartyFormBody(input(typeName), form, "HD05_"))).toEqual([]);
  });

  it("golden một giấy tờ của Đảng", () => {
    expectMatchesGolden("partyForm-giay-chung-nhan.xml", assemblePartyFormBody(input("Giấy chứng nhận"), findSmallForm("giay-chung-nhan")!, "HD05_"));
  });
});
