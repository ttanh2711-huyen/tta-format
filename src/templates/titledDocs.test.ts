import { buildCheckDocument } from "../checker/documentModel";
import { runChecks } from "../checker/engine";
import { headerFooterRules } from "../checker/ruleSets";
import { assembleDocumentBody, BODY_MARKER, type AssembleInput } from "../ooxml/documentAssembler";
import { parseStyleTable } from "../ooxml/styleResolver";
import { buildStyleSpecs } from "../ooxml/styleSpecs";
import { toStyleDefinitionsXml } from "../ooxml/styleXml";
import { extractDocModel, type ParagraphModel } from "../rules/docModel";
import { assemblePartyDocumentBody, type PartyAssembleInput } from "../standards/hd05/assembleDocument";
import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import type { StandardProfile } from "../standards/types";
import { findTemplate, templatesFor } from "./registry";

function check(profile: StandardProfile, bodyXml: string) {
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml.replace(BODY_MARKER, "")}</w:body></w:document>`;
  const styles = parseStyleTable(`<w:styles>${toStyleDefinitionsXml(buildStyleSpecs(profile))}</w:styles>`);
  const page = { widthMm: 210, heightMm: 297, marginsMm: { ...profile.pageDefaults.marginsMm }, headerDistanceMm: profile.pageDefaults.headerDistanceMm, footerDistanceMm: profile.pageDefaults.footerDistanceMm, mirrorMargins: false };
  const model = extractDocModel({ documentXml: xml });
  return { model, issues: runChecks(buildCheckDocument(model, styles, [page]), profile, headerFooterRules(profile)) };
}

const nd30Input = (typeName: string, abbr: string): AssembleInput => ({
  header: {
    parentLines: ["Sở Nông nghiệp và Môi trường"],
    issuer: "Chi cục Thủy lợi",
    sizes: { parent: 13, issuer: 13, quocHieu: 13, tieuNgu: 14, so: 13, placeDate: 13 },
    agencyLineRatio: 1 / 2,
    textWidthMm: 160,
    numberRow: { so: `Số: 05/${abbr}-CCTL`, placeDate: "Cần Thơ, ngày 05 tháng 01 năm 2026" },
  },
  title: { variant: "named", typeName, subject: "Về việc triển khai nhiệm vụ", recipients: [], sizes: { tenLoai: 14, trichYeu: 14, kinhGui: 14 }, titleLineRatio: 1 / 2, textWidthMm: 160, lineMultiple: 1.17 },
  footer: {
    recipients: ["Như trên"],
    luu: "VT",
    authority: "",
    positionLines: ["Chi cục trưởng"],
    signerName: "Mai Hiếu Hiền",
    signGapLines: 6,
    sizes: { noiNhanTieuDe: 12, noiNhanDanhSach: 11, quyenHanChucVu: 14, hoTen: 14 },
    textWidthMm: 160,
  },
});

describe("nhóm có tên loại và trích yếu (Chính quyền, Mẫu 1.4)", () => {
  const titled = templatesFor(ND30).filter((entry) => entry.family === "titled");

  it.each(titled.map((entry) => [entry.label, entry] as const))("%s: khối đầu, tên loại in hoa, trích yếu, nội dung, khối cuối; 0 lỗi", (_label, entry) => {
    const { model, issues } = check(ND30, assembleDocumentBody(nd30Input(entry.typeName, entry.abbr ?? ""), {}));
    const texts = model.blocks.filter((b): b is ParagraphModel => b.kind === "paragraph").map((p) => p.text);
    expect(texts[0]).toBe(entry.typeName.toUpperCase());
    expect(texts[1]).toBe("Về việc triển khai nhiệm vụ");
    expect(model.blocks.map((b) => b.kind)[0]).toBe("table");
    expect(issues).toEqual([]);
  });
});

const partyInput = (typeName: string): PartyAssembleInput => {
  const sizes = { theLoai: 16, trichYeu: 14, trichYeuCongVan: 12, kinhGui: 14 };
  return {
    header: { parentLines: ["Tỉnh ủy Đồng Tháp"], issuerLines: ["Ban Tổ chức"], sizes: { title: 15, parent: 14, issuer: 14, so: 14, placeDate: 14 }, numberRow: { so: "Số 05-BC/BTCTU", placeDate: "Đồng Tháp, ngày 03 tháng 02 năm 2026" }, textWidthMm: 165 },
    title: { variant: "coTenLoai", typeName, subject: "kết quả thực hiện nhiệm vụ", recipients: [], basis: [], sizes, lineExactPt: 20, textWidthMm: 165 },
    footer: {
      recipients: ["Như trên"],
      luu: "VT",
      drafter: "",
      authority: "",
      positionLines: ["Trưởng ban"],
      signerName: "Nguyễn Văn A",
      confirmation: "",
      signGapLines: 6,
      sizes: { noiNhanTu: 14, noiNhanDanhSach: 12, quyenHan: 14, chucVu: 14, hoTen: 14, kyHieuNguoiSoan: 11 },
      textWidthMm: 165,
    },
  };
};

describe("thể loại có tên loại (Đảng)", () => {
  const titled = templatesFor(HD05).filter((entry) => entry.family === "titled");

  it("có nhiều thể loại trong nhóm", () => {
    expect(titled.length).toBeGreaterThan(15);
  });

  it.each(titled.map((entry) => [entry.label, entry] as const))("%s: thể loại in hoa, 0 lỗi khối đầu, tiêu đề, khối cuối", (_label, entry) => {
    const { model, issues } = check(HD05, assemblePartyDocumentBody(partyInput(entry.typeName), {}));
    const texts = model.blocks.filter((b): b is ParagraphModel => b.kind === "paragraph").map((p) => p.text);
    expect(texts[0]).toBe(entry.typeName.toUpperCase());
    expect(issues.filter((i) => i.ruleId.startsWith("header.") || i.ruleId.startsWith("footer.") || i.ruleId.startsWith("title."))).toEqual([]);
  });

  it("tờ trình có Kính trình", () => {
    expect(findTemplate(HD05, "to-trinh")?.family).toBe("toTrinh");
  });
});
