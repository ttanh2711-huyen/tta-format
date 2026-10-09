import { buildCheckDocument } from "../checker/documentModel";
import { runChecks } from "../checker/engine";
import { headerFooterRules } from "../checker/ruleSets";
import { expectMatchesGolden } from "../ooxml/__golden__/golden";
import { BODY_MARKER, type AssembleInput } from "../ooxml/documentAssembler";
import { parseStyleTable } from "../ooxml/styleResolver";
import { buildStyleSpecs } from "../ooxml/styleSpecs";
import { toStyleDefinitionsXml } from "../ooxml/styleXml";
import { extractDocModel, type ParagraphModel } from "../rules/docModel";
import { ND30 } from "../standards/nd30/profile";
import { findTemplate } from "./registry";
import { assembleSmallFormBody, buildSmallFormBody, findSmallForm, SMALL_FORMS } from "./smallForms";

const input = (typeName: string): AssembleInput => ({
  header: {
    parentLines: ["Ủy ban nhân dân thành phố"],
    issuer: "Văn phòng",
    sizes: { parent: 13, issuer: 13, quocHieu: 13, tieuNgu: 14, so: 13, placeDate: 13 },
    agencyLineRatio: 1 / 2,
    textWidthMm: 160,
    numberRow: { so: "Số: 05/GM-VP", placeDate: "Cần Thơ, ngày 05 tháng 01 năm 2026" },
  },
  title: { variant: "named", typeName, subject: "Về việc tham dự hội nghị", recipients: [], sizes: { tenLoai: 14, trichYeu: 14, kinhGui: 14 }, titleLineRatio: 1 / 2, textWidthMm: 160, lineMultiple: 1.17 },
  footer: {
    recipients: ["Như trên"],
    luu: "VT",
    authority: "",
    positionLines: ["Chủ tịch"],
    signerName: "Nguyễn Văn A",
    signGapLines: 6,
    sizes: { noiNhanTieuDe: 12, noiNhanDanhSach: 11, quyenHanChucVu: 14, hoTen: 14 },
    textWidthMm: 160,
  },
});

function paragraphsOf(xml: string): ParagraphModel[] {
  const doc = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${xml}</w:body></w:document>`;
  return extractDocModel({ documentXml: doc }).blocks.filter((b): b is ParagraphModel => b.kind === "paragraph");
}

describe("giấy tờ và mẫu nhỏ", () => {
  it("năm mẫu, mỗi mẫu có trong registry với đúng họ mẫu và ghi rõ nguồn bố cục", () => {
    expect(SMALL_FORMS.map((form) => form.docTypeId).sort()).toEqual(["cong-dien", "giay-gioi-thieu", "giay-moi", "giay-nghi-phep", "thu-cong"]);
    for (const form of SMALL_FORMS) {
      expect(findTemplate(ND30, form.docTypeId)).toMatchObject({ family: "form", supported: true });
      expect(form.source).not.toBe("");
      expect(findSmallForm(form.docTypeId)).toBe(form);
    }
  });

  it.each(SMALL_FORMS.map((form) => [form.docTypeId, form] as const))("%s: các mục có nhãn đậm, chữ đánh dấu sau mục đầu, câu kết", (_id, form) => {
    const paragraphs = paragraphsOf(buildSmallFormBody(form, "ND30_"));
    expect(paragraphs.slice(0, form.fields.length).map((p) => p.text.replace(BODY_MARKER, "").trim())).toEqual([...form.fields]);
    expect(paragraphs[0].text).toBe(`${form.fields[0]} ${BODY_MARKER}`);
    expect(paragraphs.every((p) => p.styleId === "ND30NoiDung")).toBe(true);
    if (form.closing !== "") expect(paragraphs[paragraphs.length - 1].text).toBe(form.closing);
    expect(buildSmallFormBody(form, "ND30_").split(BODY_MARKER)).toHaveLength(2);
  });

  it.each(SMALL_FORMS.map((form) => [form.docTypeId, form] as const))("%s: văn bản đầy đủ cho 0 lỗi và khớp golden", (id, form) => {
    expectMatchesGolden(`smallForm-${id}.xml`, buildSmallFormBody(form, "ND30_"));
    const type = findTemplate(ND30, id)!;
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${assembleSmallFormBody(input(type.typeName), form, "ND30_").replace(BODY_MARKER, "")}</w:body></w:document>`;
    const styles = parseStyleTable(`<w:styles>${toStyleDefinitionsXml(buildStyleSpecs(ND30))}</w:styles>`);
    expect(runChecks(buildCheckDocument(extractDocModel({ documentXml: xml }), styles, []), ND30, headerFooterRules(ND30))).toEqual([]);
  });

  it("biến thể công văn bị từ chối", () => {
    const base = input("Giấy mời");
    expect(() => assembleSmallFormBody({ ...base, title: { ...base.title!, variant: "letter" } }, SMALL_FORMS[0], "ND30_")).toThrow();
  });
});
