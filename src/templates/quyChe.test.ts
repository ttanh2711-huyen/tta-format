import { buildCheckDocument } from "../checker/documentModel";
import { runChecks } from "../checker/engine";
import { applyFix } from "../checker/fixes";
import { headerFooterRules } from "../checker/ruleSets";
import type { CheckDocument, CheckParagraph, Issue } from "../checker/types";
import { expectMatchesGolden } from "../ooxml/__golden__/golden";
import { BODY_MARKER, type AssembleInput } from "../ooxml/documentAssembler";
import { regulationLevels } from "../ooxml/numbering";
import { parseStyleTable } from "../ooxml/styleResolver";
import { buildStyleSpecs } from "../ooxml/styleSpecs";
import { toStyleDefinitionsXml } from "../ooxml/styleXml";
import { extractDocModel, type ParagraphModel } from "../rules/docModel";
import { checkNd30Structure } from "../standards/nd30/structureChecks";
import { ND30 } from "../standards/nd30/profile";
import { assembleRegulationBody, attachedText, buildRegulationBody, regulationNumberingXml, validateRegulation, type RegulationInput } from "./quyChe";

const base: RegulationInput = { chapters: 2, articlesPerChapter: 2, withSections: false, attachedLine: "Ban hành kèm theo Quyết định số 05/QĐ-VP ngày 05 tháng 01 năm 2026", stylePrefix: "ND30_", firstLineIndentMm: 12.7, sizePt: 14 };

function paragraphsOf(xml: string): ParagraphModel[] {
  const doc = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${xml}</w:body></w:document>`;
  return extractDocModel({ documentXml: doc }).blocks.filter((b): b is ParagraphModel => b.kind === "paragraph");
}

describe("kiểm tra đầu vào", () => {
  it("hợp lệ", () => {
    expect(validateRegulation(base)).toEqual([]);
  });

  it.each([
    [{ chapters: 0 }],
    [{ chapters: 21 }],
    [{ articlesPerChapter: 0 }],
    [{ articlesPerChapter: 2.5 }],
  ])("%j bị từ chối", (patch) => {
    expect(validateRegulation({ ...base, ...patch })).not.toEqual([]);
    expect(() => buildRegulationBody({ ...base, ...patch })).toThrow();
  });

  it("dòng ban hành kèm theo được đặt trong ngoặc đơn", () => {
    expect(attachedText(" Ban hành kèm theo A ")).toBe("(Ban hành kèm theo A)");
    expect(attachedText("(Đã có ngoặc)")).toBe("(Đã có ngoặc)");
    expect(attachedText("  ")).toBe("");
  });
});

describe("cấu trúc Chương, Mục, Điều", () => {
  it("mỗi Chương: nhãn, tên; mỗi Điều: nhãn và Khoản; chữ đánh dấu chỗ nội dung ở tên Chương thứ nhất", () => {
    const xml = buildRegulationBody(base);
    const paragraphs = paragraphsOf(xml);
    expect(paragraphs[0].text).toBe("(Ban hành kèm theo Quyết định số 05/QĐ-VP ngày 05 tháng 01 năm 2026)");
    expect(paragraphs[0]).toMatchObject({ italic: true, align: "center" });
    expect(paragraphs.map((p) => p.styleId).slice(1)).toEqual([
      "ND30PhanChuong", "ND30TieuDePhanChuong", "ND30Dieu", "ND30Khoan", "ND30Dieu", "ND30Khoan",
      "ND30PhanChuong", "ND30TieuDePhanChuong", "ND30Dieu", "ND30Khoan", "ND30Dieu", "ND30Khoan",
    ]);
    expect(paragraphs[2].text).toBe(BODY_MARKER);
    expect(xml.split(BODY_MARKER)).toHaveLength(2);
  });

  it("có Mục thì mỗi Chương thêm nhãn Mục và tên Mục trước các Điều", () => {
    const styles = paragraphsOf(buildRegulationBody({ ...base, chapters: 1, withSections: true, attachedLine: "" })).map((p) => p.styleId);
    expect(styles).toEqual(["ND30PhanChuong", "ND30TieuDePhanChuong", "ND30MucTieuMuc", "ND30TieuDeMucTieuMuc", "ND30Dieu", "ND30Khoan", "ND30Dieu", "ND30Khoan"]);
  });

  it("không có dòng ban hành kèm theo khi để trống", () => {
    expect(paragraphsOf(buildRegulationBody({ ...base, attachedLine: "" }))[0].styleId).toBe("ND30PhanChuong");
  });

  it.each([
    ["khongMuc", base],
    ["comMuc", { ...base, chapters: 2, withSections: true }],
  ])("golden %s", (name, input) => {
    expectMatchesGolden(`quyChe-${name}.xml`, buildRegulationBody(input));
  });
});

describe("Khoản và Điểm dựng sẵn trong Điều", () => {
  const styles = (patch: Partial<RegulationInput>) => paragraphsOf(buildRegulationBody({ ...base, chapters: 1, attachedLine: "", ...patch })).map((p) => p.styleId).slice(2);

  it("mặc định mỗi Điều có một Khoản như trước", () => {
    expect(styles({ articlesPerChapter: 2 })).toEqual(["ND30Dieu", "ND30Khoan", "ND30Dieu", "ND30Khoan"]);
  });

  it("không Khoản thì Điều đứng một mình; nhiều Khoản và Điểm thì đủ thứ tự", () => {
    expect(styles({ articlesPerChapter: 2, khoanPerDieu: 0 })).toEqual(["ND30Dieu", "ND30Dieu"]);
    expect(styles({ articlesPerChapter: 1, khoanPerDieu: 2, diemPerKhoan: 2 })).toEqual(["ND30Dieu", "ND30Khoan", "ND30Diem", "ND30Diem", "ND30Khoan", "ND30Diem", "ND30Diem"]);
  });

  it("Điểm cần có Khoản; số ngoài 0 đến 10 bị từ chối", () => {
    expect(validateRegulation({ ...base, khoanPerDieu: 0, diemPerKhoan: 1 })).not.toEqual([]);
    expect(validateRegulation({ ...base, khoanPerDieu: 11 })).not.toEqual([]);
    expect(validateRegulation({ ...base, khoanPerDieu: 2, diemPerKhoan: 2 })).toEqual([]);
  });
});

describe("đánh số nhiều cấp", () => {
  const xml = regulationNumberingXml(base);
  const levels = xml.split("<w:lvl ").slice(1);

  it("năm cấp: Chương I, Mục 1, Điều 1., 1., a)", () => {
    expect(levels).toHaveLength(5);
    expect(xml).toContain('<w:lvlText w:val="Chương %1"/>');
    expect(xml).toContain('<w:numFmt w:val="upperRoman"/>');
    expect(xml).toContain('<w:lvlText w:val="Mục %2"/>');
    expect(xml).toContain('<w:lvlText w:val="Điều %3."/>');
    expect(xml).toContain('<w:lvlText w:val="%4."/>');
    expect(xml).toContain('<w:lvlText w:val="%5)"/>');
  });

  it("Chương và Mục đứng một mình, đậm; Khoản và Điểm không đậm", () => {
    expect(levels[0]).toContain('<w:suff w:val="nothing"/>');
    expect(levels[1]).toContain('<w:suff w:val="nothing"/>');
    expect(levels[0]).toContain("<w:b/>");
    expect(levels[3]).toContain('<w:b w:val="0"/>');
    expect(levels[4]).toContain('<w:b w:val="0"/>');
  });

  it("số Điều chạy liên tục qua Chương và Mục (không bắt đầu lại); số Khoản bắt đầu lại ở mỗi Điều", () => {
    expect(levels[2]).toContain('<w:lvlRestart w:val="0"/>');
    expect(levels[3]).not.toContain("lvlRestart");
    expect(regulationLevels(0).map((l) => l.id)).toEqual(["chuong", "muc", "dieu", "khoan", "diem"]);
  });
});

describe("kiểm tra cấu trúc", () => {
  function doc(items: Array<{ style: string; text: string; bold?: boolean }>): CheckDocument {
    const paragraphs: CheckParagraph[] = items.map((p, index) => ({ index, inTable: false, styleName: p.style, resolved: { kind: "paragraph", text: p.text, bold: p.bold } }));
    return { pages: [], paragraphs };
  }
  const run = (d: CheckDocument): Issue[] => checkNd30Structure(d, ND30);
  const ids = (issues: Issue[]) => issues.map((i) => i.ruleId);

  it("đúng thì không lỗi", () => {
    expect(run(doc([{ style: "ND30_TieuDePhanChuong", text: "QUY ĐỊNH CHUNG", bold: true }, { style: "ND30_Dieu", text: "Phạm vi điều chỉnh", bold: true }]))).toEqual([]);
  });

  it("tên Chương và tên Mục viết thường hoặc không đậm bị báo, có nút sửa", () => {
    for (const [style, rule] of [["ND30_TieuDePhanChuong", "chuong"], ["ND30_TieuDeMucTieuMuc", "muc"]] as const) {
      const d = doc([{ style, text: "Quy định chung", bold: false }]);
      expect(ids(run(d))).toEqual([`structure.${rule}.upper`, `structure.${rule}.bold`]);
      let fixed = d;
      for (const issue of run(d)) fixed = applyFix(fixed, issue.fix!);
      expect(run(fixed)).toEqual([]);
    }
  });

  it("đoạn trống không bị báo", () => {
    expect(run(doc([{ style: "ND30_TieuDePhanChuong", text: "" }, { style: "ND30_Dieu", text: "" }]))).toEqual([]);
  });

  it("Điều không đậm bị báo", () => {
    expect(ids(run(doc([{ style: "ND30_Dieu", text: "Phạm vi", bold: false }])))).toEqual(["structure.dieu.bold"]);
  });

  it("nhãn Điều gõ tay phải liên tục", () => {
    const issues = run(doc([{ style: "ND30_Dieu", text: "Điều 1. A", bold: true }, { style: "ND30_Dieu", text: "Điều 3. B", bold: true }]));
    expect(ids(issues)).toEqual(["structure.dieu.sequence"]);
    expect(issues[0].expected).toBe("Điều 2");
  });
});

describe("quy chế dựng từ trình tạo cho 0 lỗi", () => {
  const input: AssembleInput = {
    header: {
      parentLines: ["Ủy ban nhân dân thành phố"],
      issuer: "Văn phòng",
      sizes: { parent: 13, issuer: 13, quocHieu: 13, tieuNgu: 14, so: 13, placeDate: 13 },
      agencyLineRatio: 1 / 2,
      textWidthMm: 160,
      numberRow: { so: "Số: 05/QC-VP", placeDate: "Cần Thơ, ngày 05 tháng 01 năm 2026" },
    },
    title: { variant: "named", typeName: "Quy chế", subject: "Làm việc của Văn phòng", recipients: [], sizes: { tenLoai: 14, trichYeu: 14, kinhGui: 14 }, titleLineRatio: 1 / 2, textWidthMm: 160, lineMultiple: 1.17 },
    footer: {
      recipients: ["Như Điều 3"],
      luu: "VT",
      authority: "",
      positionLines: ["Chủ tịch"],
      signerName: "Nguyễn Văn A",
      signGapLines: 6,
      sizes: { noiNhanTieuDe: 12, noiNhanDanhSach: 11, quyenHanChucVu: 14, hoTen: 14 },
      textWidthMm: 160,
    },
  };

  it("khối đầu, tên loại, cấu trúc, khối cuối; kiểm tra không có lỗi", () => {
    for (const withSections of [false, true]) {
      const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${assembleRegulationBody(input, { ...base, withSections }).replace(BODY_MARKER, "")}</w:body></w:document>`;
      const styles = parseStyleTable(`<w:styles>${toStyleDefinitionsXml(buildStyleSpecs(ND30))}</w:styles>`);
      const document = buildCheckDocument(extractDocModel({ documentXml: xml }), styles, []);
      expect(runChecks(document, ND30, headerFooterRules(ND30))).toEqual([]);
    }
  });

  it("biến thể công văn bị từ chối", () => {
    expect(() => assembleRegulationBody({ ...input, title: { ...input.title!, variant: "letter" } }, base)).toThrow();
  });
});
