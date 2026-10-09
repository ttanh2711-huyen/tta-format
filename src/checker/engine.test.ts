import { assembleDocumentBody, BODY_MARKER } from "../ooxml/documentAssembler";
import { buildStyleSpecs } from "../ooxml/styleSpecs";
import { toStyleDefinitionsXml } from "../ooxml/styleXml";
import { parseStyleTable } from "../ooxml/styleResolver";
import { extractDocModel, type ParagraphModel } from "../rules/docModel";
import { assemblePartyDocumentBody } from "../standards/hd05/assembleDocument";
import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import type { StandardProfile } from "../standards/types";
import { buildCheckDocument } from "./documentModel";
import { runChecks, summarizeIssues } from "./engine";
import { applyFix, fixIndex } from "./fixes";
import type { CheckDocument, CheckParagraph, PageInfo } from "./types";

const pageOf = (profile: StandardProfile, over: Partial<PageInfo> = {}): PageInfo => ({
  widthMm: 210,
  heightMm: 297,
  marginsMm: { ...profile.pageDefaults.marginsMm },
  headerDistanceMm: profile.pageDefaults.headerDistanceMm,
  footerDistanceMm: profile.pageDefaults.footerDistanceMm,
  mirrorMargins: false,
  ...over,
});

const LONG = "Thực hiện Thông báo số 458/TB-VPUB của Văn phòng Ủy ban nhân dân thành phố về tập huấn bồi dưỡng.";

function paragraph(resolved: Partial<ParagraphModel> & { text: string }, extra: Partial<CheckParagraph> = {}): CheckParagraph {
  return { index: 0, inTable: false, resolved: { kind: "paragraph", ...resolved }, ...extra };
}

const NORMAL_ND30: Partial<ParagraphModel> = {
  align: "both",
  firstLineIndentMm: 12.7,
  spaceBeforePt: 6,
  spaceAfterPt: 6,
  lineMultiple: 1.17,
  fontName: "Times New Roman",
  sizePt: 14,
};
const NORMAL_HD05: Partial<ParagraphModel> = { ...NORMAL_ND30, firstLineIndentMm: 10, lineMultiple: undefined, lineExactPt: 20 };

function check(profile: StandardProfile, paragraphs: CheckParagraph[], pages: PageInfo[] = [pageOf(profile)]) {
  return runChecks({ pages, paragraphs: paragraphs.map((p, index) => ({ ...p, index })) }, profile);
}
const ids = (issues: ReturnType<typeof check>) => issues.map((i) => i.ruleId);

describe("trang: chuẩn Chính quyền", () => {
  it("trang đúng mặc định thì không có lỗi", () => {
    expect(check(ND30, [])).toEqual([]);
  });

  it("khổ giấy khác A4", () => {
    const issues = check(ND30, [], [pageOf(ND30, { widthMm: 215.9, heightMm: 279.4 })]);
    expect(ids(issues)).toEqual(["page.paper"]);
    expect(issues[0]).toMatchObject({ severity: "error", location: { kind: "page", section: 1 } });
    expect(issues[0].message).toContain("215,9 × 279,4");
  });

  it.each([
    ["top", { top: 15 }, "page.margin.top"],
    ["bottom", { bottom: 30 }, "page.margin.bottom"],
    ["left", { left: 25 }, "page.margin.left"],
    ["right", { right: 10 }, "page.margin.right"],
  ])("lề %s ngoài khoảng cho phép", (_name, over, ruleId) => {
    const base = pageOf(ND30);
    const issues = check(ND30, [], [{ ...base, marginsMm: { ...base.marginsMm, ...over } }]);
    expect(ids(issues)).toEqual([ruleId]);
    expect(issues[0].reference).not.toBe("");
  });

  it("lề trong khoảng cho phép, kể cả lề phải 15 mm và lề trái 35 mm, không báo", () => {
    const base = pageOf(ND30);
    expect(check(ND30, [], [{ ...base, marginsMm: { ...base.marginsMm, left: 35, right: 15 } }])).toEqual([]);
  });

  it("nhiều phần: báo đúng số thứ tự phần có lỗi", () => {
    const bad = pageOf(ND30, { marginsMm: { top: 20, bottom: 20, left: 40, right: 20 } });
    const issues = check(ND30, [], [pageOf(ND30), bad]);
    expect(issues).toHaveLength(1);
    expect(issues[0].location).toEqual({ kind: "page", section: 2 });
  });

  it("trang ngang: chuẩn Chính quyền chỉ cho phép trang ngang cho bảng, biểu nên mức cảnh báo hay lỗi theo chuẩn", () => {
    const issues = check(ND30, [], [pageOf(ND30, { widthMm: 297, heightMm: 210 })]);
    expect(ids(issues)).toEqual(["page.orientation"]);
    expect(issues[0].severity).toBe(ND30.page.landscapeAllowedForTables ? "warning" : "error");
  });
});

describe("trang: chuẩn Đảng", () => {
  it("trang đúng mặc định thì không có lỗi", () => {
    expect(check(HD05, [])).toEqual([]);
  });

  it("lề phải 20 mm đúng của chính quyền nhưng sai với Đảng", () => {
    const base = pageOf(HD05);
    expect(ids(check(HD05, [], [{ ...base, marginsMm: { ...base.marginsMm, right: 20 } }]))).toEqual(["page.margin.right"]);
  });

  it("số trang phải cách mép trên 10 mm: header 12,5 mm bị cảnh báo", () => {
    const issues = check(HD05, [], [pageOf(HD05, { headerDistanceMm: 12.5 })]);
    expect(ids(issues)).toEqual(["page.headerDistance"]);
    expect(issues[0].severity).toBe("warning");
  });

  it("chuẩn Chính quyền không báo khoảng cách header", () => {
    expect(check(ND30, [], [pageOf(ND30, { headerDistanceMm: 25 })])).toEqual([]);
  });
});

describe("đoạn nội dung: chuẩn Chính quyền", () => {
  const ok = paragraph({ text: LONG, ...NORMAL_ND30 });

  it("đoạn đúng thể thức không có lỗi", () => {
    expect(check(ND30, [ok])).toEqual([]);
  });

  it.each([
    ["phông", { fontName: "Arial" }, "font.name"],
    ["cỡ chữ quá lớn", { sizePt: 16 }, "font.size"],
    ["cỡ chữ quá nhỏ", { sizePt: 12 }, "font.size"],
    ["căn trái", { align: "left" as const }, "paragraph.align"],
    ["không thụt dòng đầu", { firstLineIndentMm: undefined }, "paragraph.firstLineIndent"],
    ["thụt sai", { firstLineIndentMm: 15 }, "paragraph.firstLineIndent"],
    ["giãn đoạn bằng 0", { spaceBeforePt: 0, spaceAfterPt: 0 }, "paragraph.spacing"],
    ["giãn dòng 2", { lineMultiple: 2 }, "paragraph.lineSpacing"],
    ["giãn dòng Exactly", { lineMultiple: undefined, lineExactPt: 20 }, "paragraph.lineSpacing"],
    ["giãn dòng tối thiểu", { lineMultiple: undefined, lineAtLeastPt: 15 }, "paragraph.lineSpacing"],
  ])("lỗi %s", (_name, over, ruleId) => {
    const issues = check(ND30, [paragraph({ text: LONG, ...NORMAL_ND30, ...over })]);
    expect(ids(issues)).toEqual([ruleId]);
    expect(issues[0]).toMatchObject({ severity: "error", location: { kind: "paragraph", index: 0 } });
    expect(issues[0].message.length).toBeGreaterThan(10);
    expect(issues[0].reference).not.toBe("");
  });

  it("giãn đoạn trước 6 sau 0 (tổng 6 pt) đạt", () => {
    expect(check(ND30, [paragraph({ text: LONG, ...NORMAL_ND30, spaceAfterPt: 0 })])).toEqual([]);
  });

  it("nhiều lỗi trong một đoạn báo đủ", () => {
    const issues = check(ND30, [paragraph({ text: LONG, ...NORMAL_ND30, fontName: "Arial", sizePt: 11, align: "left" })]);
    expect(ids(issues).sort()).toEqual(["font.name", "font.size", "paragraph.align"]);
  });

  it("không báo: đoạn trống, đoạn trong bảng, đoạn căn giữa, đoạn gạch đầu dòng thụt treo, dòng ngắn căn trái", () => {
    const bad = { fontName: "Arial", sizePt: 9, align: "center" as const };
    expect(check(ND30, [paragraph({ text: "", fontName: "Arial" })])).toEqual([]);
    expect(check(ND30, [paragraph({ text: LONG, fontName: "Arial", sizePt: 9 }, { inTable: true })])).toEqual([]);
    expect(check(ND30, [paragraph({ text: LONG, ...bad })])).toEqual([]);
    expect(check(ND30, [paragraph({ text: "- " + LONG, ...NORMAL_ND30, align: "left", firstLineIndentMm: undefined, hangingIndentMm: 4, leftIndentMm: 4 })])).toEqual([]);
    expect(check(ND30, [paragraph({ text: "Kính gửi các đơn vị", ...NORMAL_ND30, align: "left", firstLineIndentMm: undefined })])).toEqual([]);
  });

  it("phông và cỡ chữ vẫn được kiểm tra với dòng ngắn căn trái", () => {
    expect(ids(check(ND30, [paragraph({ text: "Dòng ngắn", ...NORMAL_ND30, align: "left", fontName: "Arial" })]))).toEqual(["font.name"]);
  });

  it("đoạn dùng style của chuẩn được kiểm tra theo thành phần của style", () => {
    const dieu = (sizePt: number) => paragraph({ text: "Điều 1. Phạm vi điều chỉnh", fontName: "Times New Roman", sizePt, bold: true }, { styleName: "ND30_Dieu" });
    expect(check(ND30, [dieu(14)])).toEqual([]);
    const issues = check(ND30, [dieu(18)]);
    expect(ids(issues)).toEqual(["font.size"]);
    expect(issues[0].message.toLowerCase()).toContain("điều");
  });

  it("lỗi sắp xếp: lỗi trang trước, rồi lỗi đoạn theo thứ tự", () => {
    const bad = pageOf(ND30, { marginsMm: { top: 20, bottom: 20, left: 25, right: 20 } });
    const issues = check(ND30, [paragraph({ text: LONG, ...NORMAL_ND30 }), paragraph({ text: LONG, ...NORMAL_ND30, fontName: "Arial" })], [bad]);
    expect(issues.map((i) => i.location.kind)).toEqual(["page", "paragraph"]);
    expect(issues[1].location).toMatchObject({ kind: "paragraph", index: 1 });
  });

  it("tóm tắt số lỗi và cảnh báo", () => {
    const issues = check(HD05, [], [pageOf(HD05, { headerDistanceMm: 12.5 })]);
    expect(summarizeIssues(issues)).toEqual({ errors: 0, warnings: 1 });
  });
});

describe("đoạn nội dung: chuẩn Đảng", () => {
  const ok = paragraph({ text: LONG, ...NORMAL_HD05 });

  it("đoạn đúng thể thức (Exactly 20 pt, thụt 10 mm) không có lỗi", () => {
    expect(check(HD05, [ok])).toEqual([]);
  });

  it("giãn dòng Single thay vì Exactly bị báo, nêu rõ cần Exactly", () => {
    const issues = check(HD05, [paragraph({ text: LONG, ...NORMAL_HD05, lineExactPt: undefined, lineMultiple: 1 })]);
    expect(ids(issues)).toEqual(["paragraph.lineSpacing"]);
    expect(issues[0].message).toContain("Exactly");
  });

  it.each([
    ["Exactly 25 pt", { lineExactPt: 25 }],
    ["Exactly 15 pt", { lineExactPt: 15 }],
    ["tối thiểu 20 pt", { lineExactPt: undefined, lineAtLeastPt: 20 }],
  ])("giãn dòng %s bị báo", (_name, over) => {
    expect(ids(check(HD05, [paragraph({ text: LONG, ...NORMAL_HD05, ...over })]))).toEqual(["paragraph.lineSpacing"]);
  });

  it("thụt dòng đầu 12,7 mm (của chính quyền) bị báo; 10 mm đạt", () => {
    expect(ids(check(HD05, [paragraph({ text: LONG, ...NORMAL_HD05, firstLineIndentMm: 12.7 })]))).toEqual(["paragraph.firstLineIndent"]);
  });

  it("cỡ chữ nội dung 13 pt bị báo (Đảng 14–15), 15 pt đạt", () => {
    expect(ids(check(HD05, [paragraph({ text: LONG, ...NORMAL_HD05, sizePt: 13 })]))).toEqual(["font.size"]);
    expect(check(HD05, [paragraph({ text: LONG, ...NORMAL_HD05, sizePt: 15 })])).toEqual([]);
  });
});

describe("văn bản dựng từ trình tạo cho 0 lỗi", () => {
  function docFrom(bodyXml: string, profile: StandardProfile): CheckDocument {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml.replace(BODY_MARKER, "")}</w:body></w:document>`;
    const styles = parseStyleTable(`<w:styles>${toStyleDefinitionsXml(buildStyleSpecs(profile))}</w:styles>`);
    return buildCheckDocument(extractDocModel({ documentXml: xml }), styles, [pageOf(profile)]);
  }

  const sizesND = { parent: 13, issuer: 13, quocHieu: 13, tieuNgu: 14, so: 13, placeDate: 13, subject: 13 };
  const nd30: Parameters<typeof assembleDocumentBody>[0] = {
    header: {
      parentLines: ["Sở A"],
      issuer: "Chi cục B",
      sizes: sizesND,
      agencyLineRatio: 0.5,
      textWidthMm: 160,
      numberRow: { so: "Số: 05/CCB", placeDate: "Cần Thơ, ngày 05 tháng 10 năm 2026" },
      subjectRow: { text: "V/v cử công chức" },
      recipientsRow: { label: "Kính gửi:", lines: ["Sở C"], sizePt: 14 },
    },
    footer: {
      recipients: ["Như trên"],
      luu: "VT",
      authority: "KT.",
      positionLines: ["Chi cục trưởng", "Phó Chi cục trưởng"],
      signerName: "Mai Hiếu Hiền",
      signGapLines: 6,
      sizes: { noiNhanTieuDe: 12, noiNhanDanhSach: 11, quyenHanChucVu: 14, hoTen: 14 },
      textWidthMm: 160,
    },
  };

  it("công văn chuẩn Chính quyền", () => {
    expect(runChecks(docFrom(assembleDocumentBody(nd30), ND30), ND30)).toEqual([]);
  });

  it("quyết định có tên loại chuẩn Chính quyền", () => {
    const named: typeof nd30 = {
      ...nd30,
      header: { ...nd30.header, subjectRow: undefined, recipientsRow: undefined },
      title: { variant: "named", typeName: "Quyết định", subject: "Về việc ban hành Quy chế", recipients: [], sizes: { tenLoai: 14, trichYeu: 14, kinhGui: 14 }, titleLineRatio: 0.5, textWidthMm: 160, lineMultiple: 1.17 },
    };
    expect(runChecks(docFrom(assembleDocumentBody(named), ND30), ND30)).toEqual([]);
  });

  it("văn bản Đảng có thể loại và căn cứ", () => {
    const party: Parameters<typeof assemblePartyDocumentBody>[0] = {
      header: { parentLines: ["Tỉnh ủy"], issuerLines: ["Ban Tổ chức"], sizes: { title: 15, parent: 14, issuer: 14, so: 14, placeDate: 14 }, textWidthMm: 165, numberRow: { so: "Số 05-QĐ/BTCTU", placeDate: "Hà Nội, ngày 03 tháng 02 năm 2026" } },
      title: { variant: "coTenLoai", typeName: "Quyết định", subject: "Về việc ban hành Quy chế", recipients: [], basis: ["Căn cứ Điều lệ Đảng", "Căn cứ Quy định 399-QĐ/TW"], sizes: { theLoai: 16, trichYeu: 14, trichYeuCongVan: 12, kinhGui: 14 }, lineExactPt: 20, textWidthMm: 165 },
      footer: { recipients: ["Như trên"], luu: "VT", drafter: "", authority: "K/T", positionLines: ["Trưởng ban", "Phó Trưởng ban"], signerName: "Nguyễn Văn A", confirmation: "", signGapLines: 6, sizes: { noiNhanTu: 14, noiNhanDanhSach: 12, quyenHan: 14, chucVu: 14, hoTen: 14, kyHieuNguoiSoan: 11 }, textWidthMm: 165 },
    };
    expect(runChecks(docFrom(assemblePartyDocumentBody(party), HD05), HD05)).toEqual([]);
  });

  it("văn bản chuẩn Chính quyền không bị báo lỗi theo luật Đảng ở phần chung với tài liệu rỗng đoạn", () => {
    // Tài liệu của chuẩn này kiểm theo chuẩn kia chỉ nên báo lệch về lề và cỡ; ở đây chỉ kiểm tra không ném lỗi.
    expect(() => runChecks(docFrom(assembleDocumentBody(nd30), ND30), HD05)).not.toThrow();
  });
});

describe("sửa lỗi trang và đoạn bằng cách sửa đi kèm", () => {
  const BAD: Partial<ParagraphModel> = {
    align: "left",
    firstLineIndentMm: 0,
    spaceBeforePt: 0,
    spaceAfterPt: 0,
    lineMultiple: 2,
    lineExactPt: undefined,
    fontName: "Arial",
    sizePt: 9,
  };

  it.each([
    ["Chính quyền", ND30, NORMAL_ND30],
    ["Đảng", HD05, NORMAL_HD05],
  ])("áp mọi cách sửa thì hết lỗi (%s) và không đụng đoạn khác", (_name, profile, good) => {
    const paragraphs = [paragraph({ text: LONG, ...BAD }), paragraph({ text: LONG, ...good })];
    const pages = [pageOf(profile, { widthMm: 200, heightMm: 280, marginsMm: { top: 5, bottom: 5, left: 5, right: 5 }, headerDistanceMm: 30 })];
    let doc: CheckDocument = { pages, paragraphs: paragraphs.map((q, index) => ({ ...q, index })) };
    const untouched = doc.paragraphs[1];
    const first = runChecks(doc, profile);
    expect(first.length).toBeGreaterThan(5);
    expect(first.every((i) => i.fix !== undefined)).toBe(true);
    for (const issue of first) {
      const index = fixIndex(issue.fix!);
      if (index !== undefined) expect(index).toBe(issue.location.kind === "paragraph" ? issue.location.index : -1);
      doc = applyFix(doc, issue.fix!);
    }
    expect(runChecks(doc, profile)).toEqual([]);
    expect(doc.paragraphs[1]).toEqual(untouched);
  });

  it("trang ngang không có cách sửa tự động", () => {
    const issues = runChecks({ pages: [pageOf(ND30, { widthMm: 297, heightMm: 210 })], paragraphs: [] }, ND30);
    expect(issues.filter((i) => i.ruleId === "page.orientation").every((i) => i.fix === undefined)).toBe(true);
  });
});

describe("nội dung kết thúc bằng dấu chấm (chuẩn Đảng)", () => {
  const noPeriod = LONG.replace(/\.$/, "");

  it("đoạn kết thúc bằng chữ bị cảnh báo kèm cách sửa; chuẩn Chính quyền không có quy tắc này", () => {
    const issues = check(HD05, [paragraph({ text: noPeriod, ...NORMAL_HD05 })]);
    expect(ids(issues)).toEqual(["paragraph.endsWithPeriod"]);
    expect(issues[0].severity).toBe("warning");
    const doc: CheckDocument = { pages: [pageOf(HD05)], paragraphs: [{ ...paragraph({ text: noPeriod, ...NORMAL_HD05 }), index: 0 }] };
    expect(runChecks(applyFix(doc, issues[0].fix!), HD05)).toEqual([]);
    expect(check(ND30, [paragraph({ text: noPeriod, ...NORMAL_ND30 })])).toEqual([]);
  });

  it("đoạn kết thúc bằng dấu hai chấm, chấm phẩy hoặc chấm thì không báo", () => {
    for (const end of [":", ";", "."]) expect(check(HD05, [paragraph({ text: `${noPeriod}${end}`, ...NORMAL_HD05 })])).toEqual([]);
  });
});
