import { assembleDocumentBody, BODY_MARKER, type AssembleInput } from "../ooxml/documentAssembler";
import { parseStyleTable } from "../ooxml/styleResolver";
import { extractDocModel } from "../rules/docModel";
import { assemblePartyDocumentBody, type PartyAssembleInput } from "../standards/hd05/assembleDocument";
import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import type { StandardProfile } from "../standards/types";
import { buildCheckDocument } from "./documentModel";
import { runChecks } from "./engine";
import { simulateFixAll } from "./fixAll";
import { headerFooterRules } from "./ruleSets";
import type { CheckDocument, CheckParagraph, Issue, PageInfo } from "./types";

/**
 * Checkpoint 4, phần tự động: bộ văn bản mẫu của cả hai chuẩn dựng từ trình tạo phải cho 0 lỗi; mỗi lỗi cài sẵn
 * phải được báo đúng mã quy tắc; "sửa tất cả" giả lập phải hết lỗi sửa được; tài liệu trộn hai chuẩn được cảnh báo.
 */

const ND30_INPUT: AssembleInput = {
  header: {
    parentLines: ["Sở Nông nghiệp và Môi trường", "Thành phố Cần Thơ"],
    issuer: "Chi cục Thủy lợi",
    sizes: { parent: 13, issuer: 13, quocHieu: 13, tieuNgu: 14, so: 13, placeDate: 13, subject: 13 },
    agencyLineRatio: 1 / 2,
    textWidthMm: 160,
    numberRow: { so: "Số: 05/CCTL-ĐĐPCTT", placeDate: "Cần Thơ, ngày 05 tháng 01 năm 2026" },
    subjectRow: { text: "V/v cử công chức tham dự Hội nghị tập huấn" },
    recipientsRow: { label: "Kính gửi:", lines: ["Sở Nông nghiệp và Môi trường thành phố Cần Thơ"], sizePt: 14 },
  },
  title: {
    variant: "letter",
    subject: "cử công chức tham dự Hội nghị tập huấn",
    recipients: ["Sở Nông nghiệp và Môi trường thành phố Cần Thơ"],
    sizes: { tenLoai: 14, trichYeu: 14, kinhGui: 14 },
    titleLineRatio: 1 / 2,
    textWidthMm: 160,
    lineMultiple: 1.17,
  },
  footer: {
    recipients: ["Như trên", "Chi cục trưởng (để b/c)"],
    luu: "VT, CCTLTTA",
    authority: "KT.",
    positionLines: ["Chi cục trưởng", "Phó Chi cục trưởng"],
    signerName: "Mai Hiếu Hiền",
    signGapLines: 6,
    sizes: { noiNhanTieuDe: 12, noiNhanDanhSach: 11, quyenHanChucVu: 14, hoTen: 14 },
    textWidthMm: 160,
  },
};

const PARTY_SIZES = { theLoai: 16, trichYeu: 14, trichYeuCongVan: 12, kinhGui: 14 };
const PARTY_INPUT: PartyAssembleInput = {
  header: {
    parentLines: ["Tỉnh ủy Đồng Tháp"],
    issuerLines: ["Ban Tổ chức"],
    sizes: { title: 15, parent: 14, issuer: 14, so: 14, placeDate: 14 },
    numberRow: { so: "Số 05-HD/BTCTU", placeDate: "Đồng Tháp, ngày 03 tháng 02 năm 2026" },
    textWidthMm: 165,
  },
  title: { variant: "coTenLoai", typeName: "Hướng dẫn", subject: "thực hiện quy chế làm việc", recipients: [], basis: ["Căn cứ Điều lệ Đảng"], sizes: PARTY_SIZES, lineExactPt: 20, textWidthMm: 165 },
  footer: {
    recipients: ["Như trên", "Ban Thường vụ Tỉnh ủy"],
    luu: "VT, TCTU",
    drafter: "",
    authority: "K/T",
    positionLines: ["Trưởng ban", "Phó Trưởng ban"],
    signerName: "Nguyễn Văn A",
    confirmation: "",
    signGapLines: 6,
    sizes: { noiNhanTu: 14, noiNhanDanhSach: 12, quyenHan: 14, chucVu: 14, hoTen: 14, kyHieuNguoiSoan: 11 },
    textWidthMm: 165,
  },
};

const pageOf = (profile: StandardProfile, over: Partial<PageInfo> = {}): PageInfo => ({
  widthMm: 210,
  heightMm: 297,
  marginsMm: { ...profile.pageDefaults.marginsMm },
  headerDistanceMm: profile.pageDefaults.headerDistanceMm,
  footerDistanceMm: profile.pageDefaults.footerDistanceMm,
  mirrorMargins: false,
  ...over,
});

function docFrom(bodyXml: string, profile: StandardProfile): CheckDocument {
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml.replace(BODY_MARKER, "")}</w:body></w:document>`;
  return buildCheckDocument(extractDocModel({ documentXml: xml }), parseStyleTable("<w:styles></w:styles>"), [pageOf(profile)]);
}

const SETS: Array<{ name: string; profile: StandardProfile; clean: () => CheckDocument }> = [
  { name: "Chính quyền", profile: ND30, clean: () => docFrom(assembleDocumentBody(ND30_INPUT), ND30) },
  { name: "Đảng", profile: HD05, clean: () => docFrom(assemblePartyDocumentBody(PARTY_INPUT), HD05) },
];

const check = (doc: CheckDocument, profile: StandardProfile): Issue[] => runChecks(doc, profile, headerFooterRules(profile));

/** Bản sao tài liệu với đoạn đầu tiên khớp `match` bị sửa bằng `patch`. */
function mutate(doc: CheckDocument, match: (p: CheckParagraph) => boolean, patch: Partial<CheckParagraph["resolved"]>): CheckDocument {
  let done = false;
  return {
    ...doc,
    paragraphs: doc.paragraphs.map((p) => {
      if (done || !match(p)) return p;
      done = true;
      return { ...p, resolved: { ...p.resolved, ...patch } };
    }),
  };
}

describe.each(SETS)("Checkpoint 4 – chuẩn $name", ({ profile, clean }) => {
  it("văn bản dựng từ trình tạo cho 0 lỗi", () => {
    expect(check(clean(), profile)).toEqual([]);
  });

  it("lỗi trang được báo và sửa", () => {
    const doc: CheckDocument = { ...clean(), pages: [pageOf(profile, { marginsMm: { top: 5, bottom: 5, left: 5, right: 5 }, headerDistanceMm: 30 })] };
    const codes = check(doc, profile).map((i) => i.ruleId);
    expect(codes).toEqual(expect.arrayContaining(["page.margin.top", "page.margin.left"]));
    const result = simulateFixAll(doc, (d) => check(d, profile));
    expect(result.remaining.filter((i) => i.fix)).toEqual([]);
  });

  it("tài liệu trộn tiêu đề hai chuẩn được cảnh báo", () => {
    const other = profile === ND30 ? HD05 : ND30;
    const base = clean();
    const index = base.paragraphs.length;
    const foreign: CheckParagraph = { index, inTable: true, table: { index: 0, row: 0, col: 0 }, resolved: { kind: "paragraph", text: other.identifyingTitle, bold: true, sizePt: 13 } };
    const issues = check({ ...base, paragraphs: [...base.paragraphs, foreign] }, profile);
    expect(issues.map((i) => i.ruleId)).toContain("header.mixed");
    expect(issues.find((i) => i.ruleId === "header.mixed")?.fix).toBeUndefined();
  });
});

describe("Checkpoint 4 – lỗi cài sẵn chuẩn Chính quyền", () => {
  const cases: Array<[string, (d: CheckDocument) => CheckDocument, string]> = [
    ["quốc hiệu không đậm", (d) => mutate(d, (p) => p.resolved.text === "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", { bold: false }), "header.quocHieu.bold"],
    ["số thiếu dấu hai chấm", (d) => mutate(d, (p) => p.resolved.text.startsWith("Số"), { text: "Số 5/CCTL-ĐĐPCTT" }), "header.so.format"],
    ["ngày thiếu số 0", (d) => mutate(d, (p) => p.resolved.text.includes("ngày"), { text: "Cần Thơ, ngày 5 tháng 1 năm 2026" }), "header.date.format"],
    ["chức vụ viết thường", (d) => mutate(d, (p) => p.resolved.text.startsWith("KT."), { text: "KT. Chi cục trưởng" }), "footer.chucVu.upper"],
    ["họ tên người ký không đậm", (d) => mutate(d, (p) => p.resolved.text === "Mai Hiếu Hiền", { bold: false }), "footer.hoTen.bold"],
  ];
  it.each(cases)("%s", (_name, inject, ruleId) => {
    const doc = inject(SETS[0].clean());
    expect(check(doc, ND30).map((i) => i.ruleId)).toContain(ruleId);
    const result = simulateFixAll(doc, (d) => check(d, ND30));
    expect(result.remaining.filter((i) => i.fix)).toEqual([]);
    expect(result.remaining.map((i) => i.ruleId)).not.toContain(ruleId);
  });
});

describe("Checkpoint 4 – lỗi cài sẵn chuẩn Đảng", () => {
  const cases: Array<[string, (d: CheckDocument) => CheckDocument, string]> = [
    ["tiêu đề không đậm", (d) => mutate(d, (p) => p.resolved.text === "ĐẢNG CỘNG SẢN VIỆT NAM", { bold: false }), "header.tieuDe.bold"],
    ["số có dấu hai chấm", (d) => mutate(d, (p) => p.resolved.text.startsWith("Số"), { text: "Số: 05-HD/BTCTU" }), "header.so.format"],
    ["thiếu số 0", (d) => mutate(d, (p) => p.resolved.text.startsWith("Số"), { text: "Số 5-HD/BTCTU" }), "header.so.format"],
    ["thiếu 5 dấu gạch nối", (d) => mutate(d, (p) => /^-{5}$/.test(p.resolved.text), { text: "---" }), "title.dashes.count"],
    ["\"Nơi nhận\" không gạch chân", (d) => mutate(d, (p) => p.resolved.text === "Nơi nhận:", { underline: false }), "footer.noiNhan.format"],
    ["quyền hạn kiểu Chính quyền", (d) => mutate(d, (p) => p.resolved.text.startsWith("K/T"), { text: "KT. TRƯỞNG BAN" }), "footer.quyenHan.format"],
    ["họ tên người ký không đậm", (d) => mutate(d, (p) => p.resolved.text === "Nguyễn Văn A", { bold: false }), "footer.hoTen.bold"],
  ];
  it.each(cases)("%s", (_name, inject, ruleId) => {
    const doc = inject(SETS[1].clean());
    expect(check(doc, HD05).map((i) => i.ruleId)).toContain(ruleId);
    const result = simulateFixAll(doc, (d) => check(d, HD05));
    expect(result.remaining.filter((i) => i.fix)).toEqual([]);
    expect(result.remaining.map((i) => i.ruleId)).not.toContain(ruleId);
  });

  it("dùng tiêu ngữ của Chính quyền trong văn bản Đảng", () => {
    const base = SETS[1].clean();
    const tieuNgu: CheckParagraph = { index: base.paragraphs.length, inTable: true, table: { index: 0, row: 0, col: 1 }, resolved: { kind: "paragraph", text: "Độc lập - Tự do - Hạnh phúc", bold: true, sizePt: 14 } };
    expect(check({ ...base, paragraphs: [...base.paragraphs, tieuNgu] }, HD05).map((i) => i.ruleId)).toContain("header.tieuNguChinhQuyen");
  });
});

describe("Checkpoint 4 – văn bản chuẩn này không bị báo theo luật chuẩn kia", () => {
  it("văn bản Chính quyền kiểm theo chuẩn Đảng chỉ lệch ở dấu hiệu riêng, không ném lỗi", () => {
    expect(() => check(SETS[0].clean(), HD05)).not.toThrow();
    expect(check(SETS[0].clean(), HD05).map((i) => i.ruleId)).toContain("header.otherStandard");
  });
});
