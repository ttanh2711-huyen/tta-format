import { buildCheckDocument } from "../../../checker/documentModel";
import { runChecks } from "../../../checker/engine";
import { applyFix } from "../../../checker/fixes";
import { headerFooterRules } from "../../../checker/ruleSets";
import type { CheckDocument, Issue } from "../../../checker/types";
import { expectMatchesGolden } from "../../../ooxml/__golden__/golden";
import { BODY_MARKER } from "../../../ooxml/documentAssembler";
import { parseStyleTable } from "../../../ooxml/styleResolver";
import { buildStyleSpecs } from "../../../ooxml/styleSpecs";
import { toStyleDefinitionsXml } from "../../../ooxml/styleXml";
import { extractDocModel, type ParagraphModel } from "../../../rules/docModel";
import type { PartyAssembleInput } from "../assembleDocument";
import { checkPartyStructure } from "../structureChecks";
import { partyRegulationLevels } from "../numbering";
import { HD05 } from "../profile";
import { assemblePartyRegulationBody, buildPartyRegulationBody, partyRegulationNumberingXml, validatePartyRegulation, type PartyRegulationInput } from "./regulations";

const base: PartyRegulationInput = { chapters: 2, articlesPerChapter: 2, withSections: false, top: "Chương", numbering: "roman", stylePrefix: "HD05_", firstLineIndentMm: 10 };

function paragraphsOf(xml: string): ParagraphModel[] {
  const doc = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${xml}</w:body></w:document>`;
  return extractDocModel({ documentXml: doc }).blocks.filter((b): b is ParagraphModel => b.kind === "paragraph");
}

describe("kiểm tra đầu vào", () => {
  it("hợp lệ", () => {
    expect(validatePartyRegulation(base)).toEqual([]);
  });

  it.each([[{ chapters: 0 }], [{ chapters: 21 }], [{ articlesPerChapter: 0 }], [{ articlesPerChapter: 1.5 }]])("%j bị từ chối", (patch) => {
    expect(validatePartyRegulation({ ...base, ...patch })).not.toEqual([]);
    expect(() => buildPartyRegulationBody({ ...base, ...patch })).toThrow();
  });
});

describe("cấu trúc Phần/Chương, Mục, Điều", () => {
  it("mỗi Chương: nhãn, tên, các Điều và Khoản; chữ đánh dấu ở tên Chương thứ nhất", () => {
    const xml = buildPartyRegulationBody(base);
    const styles = paragraphsOf(xml).map((p) => p.styleId);
    expect(styles).toEqual([
      "HD05PhanChuong", "HD05TenPhanChuong", "HD05Dieu", "HD05Khoan", "HD05Dieu", "HD05Khoan",
      "HD05PhanChuong", "HD05TenPhanChuong", "HD05Dieu", "HD05Khoan", "HD05Dieu", "HD05Khoan",
    ]);
    expect(paragraphsOf(xml)[1].text).toBe(BODY_MARKER);
    expect(xml.split(BODY_MARKER)).toHaveLength(2);
  });

  it("có Mục thì thêm nhãn Mục (số Ả Rập) và tên Mục", () => {
    const styles = paragraphsOf(buildPartyRegulationBody({ ...base, chapters: 1, withSections: true })).map((p) => p.styleId);
    expect(styles.slice(0, 4)).toEqual(["HD05PhanChuong", "HD05TenPhanChuong", "HD05MucSo", "HD05TenMuc"]);
  });

  it.each([
    ["chuongLaMa", base],
    ["phanChu", { ...base, top: "Phần" as const, numbering: "words" as const, withSections: true }],
  ])("golden %s", (name, input) => {
    expectMatchesGolden(`partyRegulation-${name}.xml`, buildPartyRegulationBody(input));
  });
});

describe("Khoản và Điểm dựng sẵn trong Điều (chuẩn Đảng)", () => {
  const styles = (patch: Partial<PartyRegulationInput>) => paragraphsOf(buildPartyRegulationBody({ ...base, chapters: 1, ...patch })).map((p) => p.styleId).slice(2);

  it("mặc định mỗi Điều có một Khoản; thêm Khoản và Điểm đúng thứ tự", () => {
    expect(styles({ articlesPerChapter: 1 })).toEqual(["HD05Dieu", "HD05Khoan"]);
    expect(styles({ articlesPerChapter: 1, khoanPerDieu: 2, diemPerKhoan: 1 })).toEqual(["HD05Dieu", "HD05Khoan", "HD05Diem", "HD05Khoan", "HD05Diem"]);
    expect(styles({ articlesPerChapter: 1, khoanPerDieu: 0 })).toEqual(["HD05Dieu"]);
  });

  it("Điểm cần có Khoản", () => {
    expect(validatePartyRegulation({ ...base, khoanPerDieu: 0, diemPerKhoan: 2 })).not.toEqual([]);
  });
});

describe("đánh số", () => {
  const text = (top: "Phần" | "Chương", numbering: "roman" | "words") => partyRegulationNumberingXml({ ...base, top, numbering });

  it("Chương hoặc Phần với số La Mã", () => {
    expect(text("Chương", "roman")).toContain('<w:lvlText w:val="Chương %1"/>');
    expect(text("Phần", "roman")).toContain('<w:lvlText w:val="Phần %1"/>');
    expect(text("Phần", "roman")).toContain('<w:numFmt w:val="upperRoman"/>');
  });

  it("số bằng chữ: Chương thứ …", () => {
    expect(text("Chương", "words")).toContain('<w:lvlText w:val="Chương thứ %1"/>');
    expect(text("Chương", "words")).toContain('<w:numFmt w:val="vietnameseCounting"/>');
  });

  it("Mục số Ả Rập, Điều chạy liên tục, Khoản 1., Điểm a)", () => {
    const xml = text("Chương", "roman");
    expect(xml).toContain('<w:lvlText w:val="Mục %2"/>');
    expect(xml).toContain('<w:lvlText w:val="Điều %3."/>');
    expect(xml).toContain('<w:lvlRestart w:val="0"/>');
    expect(xml).toContain('<w:lvlText w:val="%4."/>');
    expect(xml).toContain('<w:lvlText w:val="%5)"/>');
    expect(partyRegulationLevels("Chương", "roman", 0).map((l) => l.bold)).toEqual([true, true, true, false, false]);
  });

  it("thụt dòng đầu 10 mm là 567 twips", () => {
    expect(text("Chương", "roman")).toContain('w:firstLine="567"');
  });
});

describe("kiểm tra cấu trúc của chuẩn Đảng", () => {
  function doc(items: Array<{ style: string; text: string; bold?: boolean }>): CheckDocument {
    return { pages: [], paragraphs: items.map((p, index) => ({ index, inTable: false, styleName: p.style, resolved: { kind: "paragraph", text: p.text, bold: p.bold } })) };
  }
  const ids = (issues: Issue[]) => issues.map((i) => i.ruleId);

  it("đúng thì không lỗi", () => {
    expect(checkPartyStructure(doc([{ style: "HD05_TenPhanChuong", text: "NHỮNG QUY ĐỊNH CHUNG", bold: true }, { style: "HD05_Dieu", text: "Phạm vi", bold: true }]), HD05)).toEqual([]);
  });

  it("tên Phần, Chương, Mục viết thường hoặc không đậm bị báo và sửa được", () => {
    for (const [style, rule] of [["HD05_TenPhanChuong", "chuong"], ["HD05_TenMuc", "muc"]] as const) {
      const d = doc([{ style, text: "Quy định chung", bold: false }]);
      expect(ids(checkPartyStructure(d, HD05))).toEqual([`structure.${rule}.upper`, `structure.${rule}.bold`]);
      let fixed = d;
      for (const issue of checkPartyStructure(d, HD05)) fixed = applyFix(fixed, issue.fix!);
      expect(checkPartyStructure(fixed, HD05)).toEqual([]);
    }
  });

  it("Điều không đậm bị báo", () => {
    expect(ids(checkPartyStructure(doc([{ style: "HD05_Dieu", text: "Phạm vi", bold: false }]), HD05))).toEqual(["structure.dieu.bold"]);
  });
});

describe("văn bản Đảng có cấu trúc dựng từ trình tạo cho 0 lỗi", () => {
  const sizes = { theLoai: 16, trichYeu: 14, trichYeuCongVan: 12, kinhGui: 14 };
  const input: PartyAssembleInput = {
    header: { parentLines: ["Tỉnh ủy Đồng Tháp"], issuerLines: ["Ban Tổ chức"], sizes: { title: 15, parent: 14, issuer: 14, so: 14, placeDate: 14 }, numberRow: { so: "Số 05-QC/BTCTU", placeDate: "Đồng Tháp, ngày 03 tháng 02 năm 2026" }, textWidthMm: 165 },
    title: { variant: "coTenLoai", typeName: "Quy chế", subject: "làm việc của Ban Tổ chức", recipients: [], basis: ["Căn cứ Điều lệ Đảng"], sizes, lineExactPt: 20, textWidthMm: 165 },
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

  it.each([
    ["Chương", "roman", false],
    ["Phần", "words", true],
  ] as const)("%s %s mục=%s", (top, numbering, withSections) => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${assemblePartyRegulationBody(input, { ...base, top, numbering, withSections }).replace(BODY_MARKER, "")}</w:body></w:document>`;
    const styles = parseStyleTable(`<w:styles>${toStyleDefinitionsXml(buildStyleSpecs(HD05))}</w:styles>`);
    const page = { widthMm: 210, heightMm: 297, marginsMm: { ...HD05.pageDefaults.marginsMm }, headerDistanceMm: HD05.pageDefaults.headerDistanceMm, footerDistanceMm: HD05.pageDefaults.footerDistanceMm, mirrorMargins: false };
    const issues = runChecks(buildCheckDocument(extractDocModel({ documentXml: xml }), styles, [page]), HD05, headerFooterRules(HD05));
    expect(issues).toEqual([]);
  });

  it("công văn bị từ chối", () => {
    expect(() => assemblePartyRegulationBody({ ...input, title: { ...input.title!, variant: "congVan" } }, base)).toThrow();
    expect(() => assemblePartyRegulationBody({ ...input, title: undefined }, base)).toThrow();
  });
});
