import { buildCheckDocument } from "../checker/documentModel";
import { runChecks } from "../checker/engine";
import { headerFooterRules } from "../checker/ruleSets";
import { expectMatchesGolden } from "../ooxml/__golden__/golden";
import { BODY_MARKER, type AssembleInput } from "../ooxml/documentAssembler";
import { multiSignFromOptions } from "../ooxml/multiSignBlock";
import { parseStyleTable } from "../ooxml/styleResolver";
import { buildStyleSpecs } from "../ooxml/styleSpecs";
import { toStyleDefinitionsXml } from "../ooxml/styleXml";
import { extractDocModel, type ParagraphModel } from "../rules/docModel";
import { ND30 } from "../standards/nd30/profile";
import { agreementNumberingXml, assembleAgreementBody, buildAgreementBody, defaultAgreementParties, validateAgreement, type AgreementInput } from "./hopDong";
import { assembleMinutesBody, buildMinutesBody, defaultMinutesParties, MINUTES_FIELDS } from "./bienBan";

const header: Pick<AssembleInput, "header" | "title"> = {
  header: {
    parentLines: ["Ủy ban nhân dân thành phố"],
    issuer: "Văn phòng",
    sizes: { parent: 13, issuer: 13, quocHieu: 13, tieuNgu: 14, so: 13, placeDate: 13 },
    agencyLineRatio: 1 / 2,
    textWidthMm: 160,
    numberRow: { so: "Số: 05/BB-VP", placeDate: "Cần Thơ, ngày 05 tháng 01 năm 2026" },
  },
  title: { variant: "named", typeName: "Biên bản", subject: "Cuộc họp giao ban", recipients: [], sizes: { tenLoai: 14, trichYeu: 14, kinhGui: 14 }, titleLineRatio: 1 / 2, textWidthMm: 160, lineMultiple: 1.17 },
};

const agreement: AgreementInput = { parties: 2, articleCount: 3, stylePrefix: "ND30_", firstLineIndentMm: 12.7 };

function paragraphsOf(xml: string): ParagraphModel[] {
  const doc = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${xml}</w:body></w:document>`;
  return extractDocModel({ documentXml: doc }).blocks.filter((b): b is ParagraphModel => b.kind === "paragraph");
}

function issuesOf(bodyXml: string) {
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml.replace(BODY_MARKER, "")}</w:body></w:document>`;
  const styles = parseStyleTable(`<w:styles>${toStyleDefinitionsXml(buildStyleSpecs(ND30))}</w:styles>`);
  return runChecks(buildCheckDocument(extractDocModel({ documentXml: xml }), styles, []), ND30, headerFooterRules(ND30));
}

const sign = (parties: Parameters<typeof multiSignFromOptions>[2]["parties"], recipients: string[] = []) => multiSignFromOptions(ND30, {}, { parties, recipients, luu: "" });

describe("biên bản", () => {
  it("các mục đầu theo thứ tự, chữ đánh dấu sau nhãn đầu tiên", () => {
    const paragraphs = paragraphsOf(buildMinutesBody("ND30_"));
    expect(paragraphs.map((p) => p.text.replace(BODY_MARKER, "").trim())).toEqual([...MINUTES_FIELDS]);
    expect(paragraphs[0].text).toBe(`Thời gian: ${BODY_MARKER}`);
    expect(paragraphs.every((p) => p.styleId === "ND30NoiDung")).toBe(true);
    expect(buildMinutesBody("ND30_").split(BODY_MARKER)).toHaveLength(2);
  });

  it("hai bên ký mặc định là chủ tọa và thư ký", () => {
    expect(defaultMinutesParties().map((p) => p.title)).toEqual(["Chủ tọa", "Thư ký"]);
  });

  it("golden thân và văn bản đầy đủ không có Nơi nhận; kiểm tra không báo lỗi", () => {
    expectMatchesGolden("bienBan-body.xml", buildMinutesBody("ND30_"));
    const xml = assembleMinutesBody(header, "ND30_", sign(defaultMinutesParties()));
    expect(xml).not.toContain("Nơi nhận");
    expect(issuesOf(xml)).toEqual([]);
  });

  it("có nơi nhận thì kiểm tra vẫn không báo lỗi", () => {
    expect(issuesOf(assembleMinutesBody(header, "ND30_", sign(defaultMinutesParties(), ["Như trên", "Lưu hồ sơ"])))).toEqual([]);
  });

  it("biến thể công văn bị từ chối", () => {
    expect(() => assembleMinutesBody({ ...header, title: { ...header.title!, variant: "letter" } }, "ND30_", sign(defaultMinutesParties()))).toThrow();
  });
});

describe("hợp đồng, bản ghi nhớ, bản thỏa thuận", () => {
  it("đầu vào: 2 đến 4 bên, 1 đến 30 Điều", () => {
    expect(validateAgreement(agreement)).toEqual([]);
    for (const patch of [{ parties: 1 }, { parties: 5 }, { articleCount: 0 }, { articleCount: 31 }]) expect(validateAgreement({ ...agreement, ...patch })).not.toEqual([]);
    expect(() => buildAgreementBody({ ...agreement, parties: 1 })).toThrow();
  });

  it.each([2, 3, 4])("%i bên: dòng mở đầu, nhãn từng bên, câu dẫn, các Điều", (count) => {
    const paragraphs = paragraphsOf(buildAgreementBody({ ...agreement, parties: count }));
    const texts = paragraphs.map((p) => p.text.replace(BODY_MARKER, "").trim());
    expect(texts[0]).toContain("chúng tôi gồm");
    expect(texts.slice(1, 1 + count)).toEqual(["A", "B", "C", "D"].slice(0, count).map((l) => `Bên ${l}:`));
    expect(texts[1 + count]).toBe("Các bên thống nhất thỏa thuận như sau:");
    expect(paragraphs).toHaveLength(2 + count + 3);
    expect(paragraphs[1].text).toBe(`Bên A: ${BODY_MARKER}`);
  });

  it("nhãn bên in đậm, Điều đánh số tự động", () => {
    const xml = buildAgreementBody(agreement);
    expect(xml).toContain("<w:b/>");
    expect(xml).toContain('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
    expect(agreementNumberingXml(agreement)).toContain('<w:lvlText w:val="Điều %1."/>');
  });

  it.each([2, 4])("golden %i bên và văn bản đầy đủ không báo lỗi, không cần Nơi nhận", (count) => {
    const body = buildAgreementBody({ ...agreement, parties: count });
    expectMatchesGolden(`hopDong-${count}.xml`, body);
    const xml = assembleAgreementBody({ ...header, title: { ...header.title!, typeName: "Hợp đồng" } }, { ...agreement, parties: count }, sign(defaultAgreementParties(count)));
    expect(issuesOf(xml)).toEqual([]);
    expect(xml).not.toContain("Nơi nhận");
  });

  it("biến thể công văn bị từ chối", () => {
    expect(() => assembleAgreementBody({ ...header, title: { ...header.title!, variant: "letter" } }, agreement, sign(defaultAgreementParties(2)))).toThrow();
  });
});
