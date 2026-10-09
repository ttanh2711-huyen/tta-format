import { buildCheckDocument } from "../../checker/documentModel";
import { runChecks } from "../../checker/engine";
import { applyFix } from "../../checker/fixes";
import { headerFooterRules } from "../../checker/ruleSets";
import type { CheckDocument, CheckParagraph, Issue } from "../../checker/types";
import { BODY_MARKER, type AssembleInput } from "../../ooxml/documentAssembler";
import { parseStyleTable } from "../../ooxml/styleResolver";
import { buildStyleSpecs } from "../../ooxml/styleSpecs";
import { toStyleDefinitionsXml } from "../../ooxml/styleXml";
import { extractDocModel } from "../../rules/docModel";
import { assembleDecisionBody, type DecisionInput } from "../../templates/quyetDinh";
import { checkNd30Basis } from "./decisionChecks";
import { ND30 } from "./profile";

interface P {
  text: string;
  italic?: boolean;
  bold?: boolean;
  sizePt?: number;
  table?: boolean;
}

function doc(paragraphs: P[]): CheckDocument {
  const list: CheckParagraph[] = paragraphs.map((p, index) => ({
    index,
    inTable: p.table === true,
    resolved: { kind: "paragraph", text: p.text, italic: p.italic, bold: p.bold, sizePt: p.sizePt ?? 14 },
  }));
  return { pages: [], paragraphs: list };
}

const run = (d: CheckDocument): Issue[] => checkNd30Basis(d, ND30);
const ids = (issues: Issue[]) => issues.map((i) => i.ruleId);
const block = (over: Partial<P>[] = []): P[] => [
  { text: "CHỦ TỊCH ỦY BAN NHÂN DÂN", bold: true },
  { text: "Căn cứ Luật A;", italic: true, ...over[0] },
  { text: "Xét đề nghị của Giám đốc Sở B,", italic: true, ...over[1] },
  { text: "QUYẾT ĐỊNH:", bold: true, ...over[2] },
  { text: "Nội dung điều 1." },
];

describe("phần căn cứ của quyết định", () => {
  it("đúng thì không lỗi", () => {
    expect(run(doc(block()))).toEqual([]);
  });

  it("không có dòng chốt thì không kiểm (đoạn bắt đầu bằng 'Căn cứ' trong thân văn bản thường)", () => {
    expect(run(doc([{ text: "Căn cứ vào quy định trên, đơn vị thực hiện như sau." }]))).toEqual([]);
  });

  it("căn cứ trong bảng không bị xét", () => {
    expect(run(doc([{ text: "Căn cứ Luật A", table: true }, { text: "QUYẾT ĐỊNH:", bold: true }]))).toEqual([]);
  });

  it("căn cứ không nghiêng được sửa", () => {
    const d = doc(block([{ italic: false }]));
    expect(ids(run(d))).toEqual(["basis.italic"]);
    expect(run(applyFix(d, run(d)[0].fix!)).length).toBe(0);
  });

  it("căn cứ sai cỡ được kẹp vào 13 đến 14", () => {
    const d = doc(block([{ sizePt: 16 }]));
    expect(ids(run(d))).toEqual(["basis.size"]);
    expect(run(applyFix(d, run(d)[0].fix!))).toEqual([]);
  });

  it("căn cứ giữa kết thúc không bằng chấm phẩy, căn cứ cuối không bằng phẩy", () => {
    for (const [patch, rule] of [
      [[{ text: "Căn cứ Luật A." }], "basis.punctuation"],
      [[{}, { text: "Xét đề nghị của Giám đốc Sở B;" }], "basis.punctuation"],
      [[{}, { text: "Xét đề nghị của Giám đốc Sở B" }], "basis.punctuation"],
    ] as const) {
      const d = doc(block([...patch]));
      expect(ids(run(d))).toEqual([rule]);
      expect(run(applyFix(d, run(d)[0].fix!))).toEqual([]);
    }
  });

  it("dòng chốt không đậm", () => {
    const d = doc(block([{}, {}, { bold: false }]));
    expect(ids(run(d))).toEqual(["basis.verdict.bold"]);
    expect(run(applyFix(d, run(d)[0].fix!))).toEqual([]);
  });

  it("QUYẾT NGHỊ cũng là dòng chốt", () => {
    const d = doc(block([{}, {}, { text: "QUYẾT NGHỊ:", bold: false }]));
    expect(ids(run(d))).toEqual(["basis.verdict.bold"]);
  });
});

describe("quyết định dựng từ trình tạo cho 0 lỗi", () => {
  const decision: DecisionInput = {
    kind: "quyetDinhTrucTiep",
    authority: "Chủ tịch Ủy ban nhân dân thành phố",
    basis: ["Căn cứ Luật Tổ chức chính quyền địa phương", "Xét đề nghị của Giám đốc Sở Nội vụ"],
    articleCount: 3,
    stylePrefix: "ND30_",
    sizes: { authority: 14, basis: 14, verdict: 14 },
    firstLineIndentMm: 12.7,
  };
  const input: AssembleInput = {
    header: {
      parentLines: ["Ủy ban nhân dân thành phố"],
      issuer: "Văn phòng",
      sizes: { parent: 13, issuer: 13, quocHieu: 13, tieuNgu: 14, so: 13, placeDate: 13 },
      agencyLineRatio: 1 / 2,
      textWidthMm: 160,
      numberRow: { so: "Số: 05/QĐ-VP", placeDate: "Cần Thơ, ngày 05 tháng 01 năm 2026" },
    },
    title: { variant: "named", typeName: "Quyết định", subject: "Về việc ban hành Quy chế", recipients: [], sizes: { tenLoai: 14, trichYeu: 14, kinhGui: 14 }, titleLineRatio: 1 / 2, textWidthMm: 160, lineMultiple: 1.17 },
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

  it.each(["nghiQuyet", "quyetDinhTrucTiep", "quyetDinhGianTiep"] as const)("%s", (kind) => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${assembleDecisionBody(input, { ...decision, kind }).replace(BODY_MARKER, "")}</w:body></w:document>`;
    const document = buildCheckDocument(extractDocModel({ documentXml: xml }), parseStyleTable(`<w:styles>${toStyleDefinitionsXml(buildStyleSpecs(ND30))}</w:styles>`), []);
    expect(runChecks(document, ND30, headerFooterRules(ND30))).toEqual([]);
  });
});
