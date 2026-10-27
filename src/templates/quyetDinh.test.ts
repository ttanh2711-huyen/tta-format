import { expectMatchesGolden } from "../ooxml/__golden__/golden";
import { BODY_MARKER, type AssembleInput } from "../ooxml/documentAssembler";
import { legalNumberingXml } from "../ooxml/numbering";
import { extractDocModel, type ParagraphModel } from "../rules/docModel";
import { assembleDecisionBody, basisLines, buildDecisionBody, decisionKindInfo, decisionNumberingXml, DECISION_KINDS, INDIRECT_LEAD, validateDecision, type DecisionInput } from "./quyetDinh";

const base: DecisionInput = {
  kind: "quyetDinhTrucTiep",
  authority: "Chủ tịch Ủy ban nhân dân thành phố Cần Thơ",
  basis: ["Căn cứ Luật Tổ chức chính quyền địa phương ngày 19 tháng 6 năm 2015.", "Căn cứ Nghị định số 30/2020/NĐ-CP;", "Xét đề nghị của Giám đốc Sở Nội vụ"],
  articleCount: 3,
  stylePrefix: "ND30_",
  sizes: { authority: 14, basis: 14, verdict: 14 },
  firstLineIndentMm: 12.7,
};

function paragraphsOf(xml: string): ParagraphModel[] {
  const doc = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${xml}</w:body></w:document>`;
  return extractDocModel({ documentXml: doc }).blocks.filter((b): b is ParagraphModel => b.kind === "paragraph");
}

describe("căn cứ", () => {
  it("mỗi căn cứ cuối dòng bằng dấu chấm phẩy, căn cứ cuối bằng dấu chấm (Phụ lục I, II.6); dấu cũ được thay", () => {
    expect(basisLines(base.basis)).toEqual([
      "Căn cứ Luật Tổ chức chính quyền địa phương ngày 19 tháng 6 năm 2015;",
      "Căn cứ Nghị định số 30/2020/NĐ-CP;",
      "Xét đề nghị của Giám đốc Sở Nội vụ.",
    ]);
  });

  it("bỏ dòng trống và gộp khoảng trắng", () => {
    expect(basisLines(["  Căn cứ   A  ", "", "Căn cứ B;"])).toEqual(["Căn cứ A;", "Căn cứ B."]);
  });

  it("một căn cứ duy nhất kết thúc bằng dấu chấm", () => {
    expect(basisLines(["Căn cứ A"])).toEqual(["Căn cứ A."]);
  });
});

describe("kiểm tra đầu vào", () => {
  it("hợp lệ thì không lỗi", () => {
    expect(validateDecision(base)).toEqual([]);
  });

  it.each([
    ["thiếu thẩm quyền", { authority: "  " }],
    ["không căn cứ", { basis: ["", " "] }],
    ["số Điều bằng 0", { articleCount: 0 }],
    ["số Điều quá lớn", { articleCount: 31 }],
    ["số Điều lẻ", { articleCount: 2.5 }],
  ])("%s bị từ chối", (_name, patch) => {
    expect(validateDecision({ ...base, ...patch })).not.toEqual([]);
    expect(() => buildDecisionBody({ ...base, ...patch })).toThrow();
  });
});

describe("thân văn bản", () => {
  const paragraphs = paragraphsOf(buildDecisionBody(base));

  it("thứ tự: thẩm quyền, căn cứ, dòng chốt, ba Điều", () => {
    expect(paragraphs.map((p) => p.text)).toEqual([
      "CHỦ TỊCH ỦY BAN NHÂN DÂN THÀNH PHỐ CẦN THƠ",
      "Căn cứ Luật Tổ chức chính quyền địa phương ngày 19 tháng 6 năm 2015;",
      "Căn cứ Nghị định số 30/2020/NĐ-CP;",
      "Xét đề nghị của Giám đốc Sở Nội vụ.",
      "QUYẾT ĐỊNH:",
      BODY_MARKER,
      "",
      "",
    ]);
  });

  it("thẩm quyền và dòng chốt đậm, canh giữa; căn cứ nghiêng, đều hai bên, cỡ chữ 14", () => {
    expect(paragraphs[0]).toMatchObject({ bold: true, align: "center", sizePt: 14 });
    expect(paragraphs[4]).toMatchObject({ bold: true, align: "center", sizePt: 14 });
    for (const p of paragraphs.slice(1, 4)) expect(p).toMatchObject({ italic: true, align: "both", sizePt: 14, styleId: "ND30CanCu" });
  });

  it("các Điều dùng style nội dung và có chữ đánh dấu chỗ nội dung ở Điều 1 duy nhất", () => {
    for (const p of paragraphs.slice(5)) expect(p.styleId).toBe("ND30NoiDung");
    expect(buildDecisionBody(base).split(BODY_MARKER)).toHaveLength(2);
  });

  it("gián tiếp: Điều 1 mở bằng câu ban hành kèm theo, con trỏ ngay sau", () => {
    const xml = buildDecisionBody({ ...base, kind: "quyetDinhGianTiep" });
    expect(paragraphsOf(xml)[5].text).toBe(`${INDIRECT_LEAD}${BODY_MARKER}`);
  });

  it("nghị quyết chốt bằng QUYẾT NGHỊ:", () => {
    expect(paragraphsOf(buildDecisionBody({ ...base, kind: "nghiQuyet" })).map((p) => p.text)).toContain("QUYẾT NGHỊ:");
    expect(decisionKindInfo("nghiQuyet").template).toBe("Mẫu 1.1");
  });

  it("ba loại đều có thông tin", () => {
    expect(DECISION_KINDS.map((k) => k.template)).toEqual(["Mẫu 1.1", "Mẫu 1.2", "Mẫu 1.3"]);
  });

  it.each(["nghiQuyet", "quyetDinhTrucTiep", "quyetDinhGianTiep"] as const)("golden %s", (kind) => {
    expectMatchesGolden(`quyetDinh-${kind}.xml`, buildDecisionBody({ ...base, kind }));
  });
});

describe("Khoản và Điểm dựng sẵn trong Điều", () => {
  const styles = (input: Partial<DecisionInput>) => paragraphsOf(buildDecisionBody({ ...base, ...input })).slice(5).map((p) => p.styleId);

  it("không nêu số Khoản, Điểm thì mỗi Điều chỉ có một đoạn như trước", () => {
    expect(styles({})).toEqual(["ND30NoiDung", "ND30NoiDung", "ND30NoiDung"]);
  });

  it("mỗi Điều có đủ Khoản, mỗi Khoản có đủ Điểm, đúng thứ tự", () => {
    const one = ["ND30NoiDung", "ND30Khoan", "ND30Diem", "ND30Diem", "ND30Khoan", "ND30Diem", "ND30Diem"];
    expect(styles({ articleCount: 2, khoanPerDieu: 2, diemPerKhoan: 2 })).toEqual([...one, ...one]);
  });

  it("chỉ Khoản, không Điểm", () => {
    expect(styles({ articleCount: 1, khoanPerDieu: 3 })).toEqual(["ND30NoiDung", "ND30Khoan", "ND30Khoan", "ND30Khoan"]);
  });

  it("Khoản đánh số tự động ở cấp 1; Điểm có nhãn gõ sẵn theo chữ cái tiếng Việt (a, b, c, d, đ…); chữ đánh dấu vẫn chỉ ở Điều 1", () => {
    const xml = buildDecisionBody({ ...base, articleCount: 1, khoanPerDieu: 1, diemPerKhoan: 6 });
    expect(xml).toContain('<w:ilvl w:val="1"/>');
    expect(xml).not.toContain('<w:ilvl w:val="2"/>');
    expect([...xml.matchAll(/<w:t xml:space="preserve">([a-zđ]\)) <\/w:t>/g)].map((m) => m[1])).toEqual(["a)", "b)", "c)", "d)", "đ)", "e)"]);
    expect(xml.split(BODY_MARKER)).toHaveLength(2);
  });

  it.each([
    ["Khoản âm", { khoanPerDieu: -1 }],
    ["Khoản quá 10", { khoanPerDieu: 11 }],
    ["Điểm quá 10", { khoanPerDieu: 1, diemPerKhoan: 11 }],
    ["Điểm lẻ", { khoanPerDieu: 1, diemPerKhoan: 1.5 }],
    ["Điểm khi không có Khoản", { khoanPerDieu: 0, diemPerKhoan: 2 }],
  ])("%s bị từ chối", (_name, patch) => {
    expect(validateDecision({ ...base, ...patch })).not.toEqual([]);
    expect(() => buildDecisionBody({ ...base, ...patch })).toThrow();
  });
});

describe("đánh số Điều, Khoản, Điểm", () => {
  const xml = decisionNumberingXml(base);

  it("ba cấp: Điều N. (đậm), N., a)", () => {
    expect(xml).toContain('<w:lvlText w:val="Điều %1."/>');
    expect(xml).toContain('<w:lvlText w:val="%2."/>');
    expect(xml).toContain('<w:lvlText w:val="%3)"/>');
    expect(xml).toContain('<w:numFmt w:val="lowerLetter"/>');
    const levels = xml.split("<w:lvl ").slice(1);
    expect(levels).toHaveLength(3);
    expect(levels[0]).toContain("<w:b/>");
    expect(levels[1]).toContain('<w:b w:val="0"/>');
  });

  it("thụt dòng đầu theo tuỳ chọn (12,7 mm là 720 twips)", () => {
    expect(xml).toContain('w:firstLine="720"');
    expect(legalNumberingXml(567)).toContain('w:firstLine="567"');
  });

  it("đoạn của Điều trỏ đúng danh sách và cấp 0", () => {
    expect(buildDecisionBody(base)).toContain('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
  });
});

describe("văn bản hoàn chỉnh", () => {
  const input: AssembleInput = {
    header: {
      parentLines: ["Ủy ban nhân dân thành phố Cần Thơ"],
      issuer: "Văn phòng",
      sizes: { parent: 13, issuer: 13, quocHieu: 13, tieuNgu: 14, so: 13, placeDate: 13 },
      agencyLineRatio: 1 / 2,
      textWidthMm: 160,
      numberRow: { so: "Số: 05/QĐ-VP", placeDate: "Cần Thơ, ngày 05 tháng 01 năm 2026" },
    },
    title: { variant: "named", typeName: "Quyết định", subject: "Về việc ban hành Quy chế làm việc", recipients: [], sizes: { tenLoai: 14, trichYeu: 14, kinhGui: 14 }, titleLineRatio: 1 / 2, textWidthMm: 160, lineMultiple: 1.17 },
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

  it("khối đầu, tên loại, thân quyết định, khối cuối theo thứ tự", () => {
    const doc = extractDocModel({ documentXml: `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${assembleDecisionBody(input, base)}</w:body></w:document>` });
    const kinds = doc.blocks.map((b) => b.kind);
    expect(kinds[0]).toBe("table");
    expect(kinds[kinds.length - 2]).toBe("table");
    const texts = doc.blocks.filter((b): b is ParagraphModel => b.kind === "paragraph").map((p) => p.text);
    expect(texts.indexOf("QUYẾT ĐỊNH")).toBeLessThan(texts.indexOf("QUYẾT ĐỊNH:"));
    expect(texts.indexOf("QUYẾT ĐỊNH:")).toBeLessThan(texts.indexOf(BODY_MARKER));
  });

  it("biến thể công văn bị từ chối", () => {
    expect(() => assembleDecisionBody({ ...input, title: { ...input.title!, variant: "letter" } }, base)).toThrow();
  });
});
