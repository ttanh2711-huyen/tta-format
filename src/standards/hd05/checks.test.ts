import { buildCheckDocument } from "../../checker/documentModel";
import { runChecks } from "../../checker/engine";
import { applyFix, fixIndex } from "../../checker/fixes";
import { checkMixedStandards } from "../../checker/mixedStandards";
import { headerFooterRules } from "../../checker/ruleSets";
import type { CheckDocument, CheckParagraph, Issue } from "../../checker/types";
import { BODY_MARKER } from "../../ooxml/documentAssembler";
import { parseStyleTable } from "../../ooxml/styleResolver";
import { extractDocModel } from "../../rules/docModel";
import { ND30 } from "../nd30/profile";
import { assemblePartyDocumentBody, type PartyAssembleInput } from "./assembleDocument";
import { checkPartyHeaderFooter } from "./checks";
import { HD05 } from "./profile";

interface P {
  text: string;
  col?: number;
  row?: number;
  body?: boolean;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  sizePt?: number;
  align?: "center" | "left";
}

/** Tài liệu một bảng (đoạn mặc định nằm trong bảng, ô theo hàng và cột) hoặc đoạn thân; chỉ số lần lượt từ 0. */
function doc(paragraphs: P[]): CheckDocument {
  const list: CheckParagraph[] = paragraphs.map((p, index) => ({
    index,
    inTable: !p.body,
    table: p.body ? undefined : { index: 0, row: p.row ?? 0, col: p.col ?? 0 },
    resolved: { kind: "paragraph", text: p.text, bold: p.bold, italic: p.italic, underline: p.underline, sizePt: p.sizePt ?? 14, align: p.align },
  }));
  return { pages: [], paragraphs: list };
}

const run = (d: CheckDocument): Issue[] => checkPartyHeaderFooter(d, HD05);
const ids = (issues: Issue[]) => issues.map((i) => i.ruleId);
const TITLE = "ĐẢNG CỘNG SẢN VIỆT NAM";

/** Áp cách sửa của lỗi `ruleId`: lỗi đó biến mất và không đoạn nào khác thay đổi. */
function expectFixWorks(d: CheckDocument, ruleId: string): void {
  const issue = run(d).find((i) => i.ruleId === ruleId);
  expect(issue?.fix).toBeDefined();
  const index = fixIndex(issue!.fix!);
  const fixed = applyFix(d, issue!.fix!);
  expect(run(fixed).filter((i) => i.ruleId === ruleId && i.location.kind === "paragraph" && i.location.index === index)).toEqual([]);
  fixed.paragraphs.forEach((p, i) => {
    if (i !== index) expect(p).toBe(d.paragraphs[i]);
  });
}

describe("văn bản dựng từ trình tạo cho 0 lỗi", () => {
  const sizes = { theLoai: 16, trichYeu: 14, trichYeuCongVan: 12, kinhGui: 14 };
  const footer: PartyAssembleInput["footer"] = {
    recipients: ["Như trên", "Ban Thường vụ Tỉnh ủy"],
    luu: "VT, TCTU",
    drafter: "NVA 10b",
    authority: "K/T",
    positionLines: ["Trưởng ban", "Phó Trưởng ban"],
    signerName: "Nguyễn Văn A",
    confirmation: "Trần Văn B, Bí thư",
    signGapLines: 6,
    sizes: { noiNhanTu: 14, noiNhanDanhSach: 12, quyenHan: 14, chucVu: 14, hoTen: 14, kyHieuNguoiSoan: 11 },
    textWidthMm: 165,
  };
  const header: PartyAssembleInput["header"] = {
    parentLines: ["Tỉnh ủy Đồng Tháp"],
    issuerLines: ["Ban Tổ chức"],
    sizes: { title: 15, parent: 14, issuer: 14, so: 14, placeDate: 14 },
    numberRow: { so: "Số 05-HD/BTCTU", placeDate: "Đồng Tháp, ngày 03 tháng 02 năm 2026" },
    textWidthMm: 165,
  };
  const base = { sizes, lineExactPt: 20, textWidthMm: 165 };

  function check(input: PartyAssembleInput): Issue[] {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${assemblePartyDocumentBody(input).replace(BODY_MARKER, "")}</w:body></w:document>`;
    const document = buildCheckDocument(extractDocModel({ documentXml: xml }), parseStyleTable("<w:styles></w:styles>"), []);
    return runChecks(document, HD05, headerFooterRules(HD05)).filter((i) => i.ruleId.startsWith("header.") || i.ruleId.startsWith("footer.") || i.ruleId.startsWith("title."));
  }

  it("công văn", () => {
    const input: PartyAssembleInput = {
      header: { ...header, subjectRow: { text: "về việc cử cán bộ", sizePt: 12 }, recipientsRow: { label: "Kính gửi:", italicLabel: true, lines: ["Ban Thường vụ Tỉnh ủy"], sizePt: 14 } },
      title: { ...base, variant: "congVan", subject: "cử cán bộ", recipients: ["Ban Thường vụ Tỉnh ủy"], basis: [] },
      footer,
    };
    expect(check(input)).toEqual([]);
  });

  it("văn bản có tên thể loại", () => {
    const input: PartyAssembleInput = {
      header,
      title: { ...base, variant: "coTenLoai", typeName: "Hướng dẫn", subject: "thực hiện quy chế làm việc", recipients: [], basis: ["Căn cứ Điều lệ Đảng", "Căn cứ Quy định số 24-QĐ/TW"] },
      footer,
    };
    expect(check(input)).toEqual([]);
  });

  it("tờ trình, nhiều nơi nhận, không quyền hạn", () => {
    const input: PartyAssembleInput = {
      header,
      title: { ...base, variant: "toTrinh", typeName: "Tờ trình", subject: "xin ý kiến về đề án", recipients: ["Ban Thường vụ Tỉnh ủy", "Ban Bí thư"], basis: [] },
      footer: { ...footer, authority: "", positionLines: ["Trưởng ban"], confirmation: "", drafter: "" },
    };
    expect(check(input)).toEqual([]);
  });
});

describe("tiêu đề và cơ quan ban hành", () => {
  it("tiêu đề đúng, đậm, cỡ 15 thì không lỗi", () => {
    expect(run(doc([{ text: TITLE, col: 1, bold: true, sizePt: 15 }]))).toEqual([]);
  });

  it("tiêu đề không đậm, sai cỡ, sai chữ", () => {
    expect(ids(run(doc([{ text: TITLE, col: 1, sizePt: 15 }])))).toEqual(["header.tieuDe.bold"]);
    expect(ids(run(doc([{ text: TITLE, col: 1, bold: true, sizePt: 13 }])))).toEqual(["header.tieuDe.size"]);
    expect(ids(run(doc([{ text: "ĐẢNG CỘNG SẢN VIỆT NAN", col: 1, bold: true, sizePt: 15 }])))).toEqual(["header.tieuDe.text"]);
    expectFixWorks(doc([{ text: TITLE, col: 1, sizePt: 15 }]), "header.tieuDe.bold");
    expectFixWorks(doc([{ text: TITLE, col: 1, bold: true, sizePt: 13 }]), "header.tieuDe.size");
    expectFixWorks(doc([{ text: "ĐẢNG CỘNG SẢN VIỆT NAN", col: 1, bold: true, sizePt: 15 }]), "header.tieuDe.text");
  });

  it("dùng tiêu ngữ của Chính quyền bị báo, không đổi chữ tự động", () => {
    const issues = run(doc([{ text: "Độc lập - Tự do - Hạnh phúc", col: 1, bold: true }]));
    expect(ids(issues)).toContain("header.tieuNguChinhQuyen");
    expect(issues.find((i) => i.ruleId === "header.tieuNguChinhQuyen")?.fix).toBeUndefined();
  });

  const left = (over: Partial<P>[] = []): P[] => [
    { text: TITLE, col: 1, bold: true, sizePt: 15 },
    { text: "TỈNH ỦY ĐỒNG THÁP", col: 0 },
    { text: "BAN TỔ CHỨC", col: 0, bold: true },
    { text: "*", col: 0 },
    ...over.map((o) => ({ text: "", ...o })),
  ];

  it("cơ quan ban hành đúng thì không lỗi", () => {
    expect(run(doc(left()))).toEqual([]);
  });

  it("ô trái trống (mất cả tên cơ quan và dấu *) bị báo, không có cách sửa tự động", () => {
    const issues = run(doc([{ text: TITLE, col: 1, bold: true, sizePt: 15 }, { text: "", col: 0 }]));
    expect(ids(issues)).toEqual(["header.coQuanBanHanh.missing"]);
    expect(issues[0].fix).toBeUndefined();
  });

  it("thiếu dấu *", () => {
    const d = doc(left().filter((p) => p.text !== "*"));
    expect(ids(run(d))).toEqual(["header.star"]);
  });

  it("cơ quan ban hành không đậm, viết thường, sai cỡ", () => {
    const make = (patch: Partial<P>) => doc(left().map((p) => (p.text === "BAN TỔ CHỨC" || p.text === "Ban Tổ chức" ? { ...p, ...patch } : p)));
    expect(ids(run(make({ bold: false })))).toEqual(["header.coQuanBanHanh.bold"]);
    expect(ids(run(make({ text: "Ban Tổ chức" })))).toEqual(["header.coQuanBanHanh.upper"]);
    expect(ids(run(make({ sizePt: 12 })))).toEqual(["header.coQuanBanHanh.size"]);
    expectFixWorks(make({ bold: false }), "header.coQuanBanHanh.bold");
    expectFixWorks(make({ text: "Ban Tổ chức" }), "header.coQuanBanHanh.upper");
    expectFixWorks(make({ sizePt: 12 }), "header.coQuanBanHanh.size");
  });
});

describe("số và ký hiệu", () => {
  const so = (text: string) => doc([{ text, row: 1 }]);

  it.each(["Số 05-HD/VPTW", "Số 123-QĐ/TU", "Số -HD/VPTW", "Số 05-CV/BTCTU-VPTU"])("%s đúng", (text) => {
    expect(run(so(text))).toEqual([]);
  });

  it.each([
    ["có dấu hai chấm", "Số: 05-HD/VPTW", "Số 05-HD/VPTW"],
    ["thiếu số 0", "Số 5-HD/VPTW", "Số 05-HD/VPTW"],
    ["thừa khoảng trắng", "Số 05 - HD / VPTW", "Số 05-HD/VPTW"],
    ["hai lỗi cùng lúc", "Số: 5-HD/VPTW", "Số 05-HD/VPTW"],
    ["số trống có dấu hai chấm", "Số: -HD/VPTW", "Số -HD/VPTW"],
  ])("%s bị báo kèm cách sửa", (_name, text, expected) => {
    const issues = run(so(text));
    expect(ids(issues)).toEqual(["header.so.format"]);
    expect(issues[0].expected).toBe(expected);
    expectFixWorks(so(text), "header.so.format");
  });

  it("số ghi kiểu Chính quyền bị báo nhưng không tự sửa", () => {
    const issues = run(so("Số: 05/QĐ-UBND"));
    expect(ids(issues)).toEqual(["header.so.format"]);
    expect(issues[0].fix).toBeUndefined();
  });
});

describe("ngày tháng và Kính gửi", () => {
  it("ngày đúng, nghiêng thì không lỗi", () => {
    expect(run(doc([{ text: "Đồng Tháp, ngày 05 tháng 01 năm 2026", col: 1, italic: true }]))).toEqual([]);
  });

  it("ngày không nghiêng, ngày thiếu số 0, ngày không tồn tại", () => {
    expect(ids(run(doc([{ text: "Đồng Tháp, ngày 05 tháng 01 năm 2026", col: 1 }])))).toEqual(["header.date.italic"]);
    expect(ids(run(doc([{ text: "Đồng Tháp, ngày 5 tháng 1 năm 2026", col: 1, italic: true }])))).toEqual(["header.date.format"]);
    expect(ids(run(doc([{ text: "Đồng Tháp, ngày 31 tháng 4 năm 2026", col: 1, italic: true }])))).toEqual(["header.date.invalid"]);
    expectFixWorks(doc([{ text: "Đồng Tháp, ngày 05 tháng 01 năm 2026", col: 1 }]), "header.date.italic");
    expectFixWorks(doc([{ text: "Đồng Tháp, ngày 5 tháng 1 năm 2026", col: 1, italic: true }]), "header.date.format");
  });

  it("nhãn Kính gửi, Kính trình phải nghiêng", () => {
    expect(run(doc([{ text: "Kính gửi:", row: 2, italic: true }]))).toEqual([]);
    expect(ids(run(doc([{ text: "Kính gửi:", row: 2 }])))).toEqual(["header.kinhGui.italic"]);
    expect(ids(run(doc([{ text: "Kính trình:", row: 2 }])))).toEqual(["header.kinhGui.italic"]);
    expectFixWorks(doc([{ text: "Kính gửi:", row: 2 }]), "header.kinhGui.italic");
  });
});

describe("thể loại, trích yếu, 5 dấu gạch nối", () => {
  const block = (over: Partial<Record<"type" | "subject" | "dashes", Partial<P>>> = {}): P[] => [
    { text: "HƯỚNG DẪN", body: true, align: "center", bold: true, sizePt: 16, ...over.type },
    { text: "thực hiện quy chế làm việc", body: true, align: "center", bold: true, ...over.subject },
    { text: "-----", body: true, align: "center", ...over.dashes },
    { text: "Căn cứ Điều lệ Đảng;", body: true, align: "left" },
  ];

  it("đúng thì không lỗi", () => {
    expect(run(doc(block()))).toEqual([]);
  });

  it("công văn không có thể loại thì không báo thiếu gạch nối", () => {
    expect(run(doc([{ text: "Căn cứ Điều lệ Đảng;", body: true, align: "left" }]))).toEqual([]);
  });

  it("thể loại sai cỡ, không đậm", () => {
    expect(ids(run(doc(block({ type: { sizePt: 14 } }))))).toEqual(["title.theLoai.size"]);
    expect(ids(run(doc(block({ type: { bold: false } }))))).toEqual(["title.theLoai.bold"]);
    expectFixWorks(doc(block({ type: { sizePt: 14 } })), "title.theLoai.size");
    expectFixWorks(doc(block({ type: { bold: false } })), "title.theLoai.bold");
  });

  it("trích yếu không đậm, sai cỡ", () => {
    expect(ids(run(doc(block({ subject: { bold: false } }))))).toEqual(["title.trichYeu.bold"]);
    expect(ids(run(doc(block({ subject: { sizePt: 12 } }))))).toEqual(["title.trichYeu.size"]);
    expectFixWorks(doc(block({ subject: { bold: false } })), "title.trichYeu.bold");
    expectFixWorks(doc(block({ subject: { sizePt: 12 } })), "title.trichYeu.size");
  });

  it("dòng gạch nối sai số lượng được sửa, thiếu dòng chỉ báo", () => {
    for (const text of ["---", "-------"]) {
      expect(ids(run(doc(block({ dashes: { text } }))))).toEqual(["title.dashes.count"]);
      expectFixWorks(doc(block({ dashes: { text } })), "title.dashes.count");
    }
    const missing = block().filter((p) => p.text !== "-----");
    const issues = run(doc(missing));
    expect(ids(issues)).toEqual(["title.dashes.missing"]);
    expect(issues[0].fix).toMatchObject({ kind: "insertAfter", text: "-----" });
  });
});

describe("khối cuối", () => {
  const footer = (over: Record<string, Partial<P>> = {}): P[] =>
    [
      { key: "nn", text: "Nơi nhận:", underline: true },
      { key: "a", text: "- Như trên;" , sizePt: 12 },
      { key: "b", text: "- Lưu: VT, TCTU.", sizePt: 12 },
      { key: "d", text: "NVA 10b", sizePt: 11 },
      { key: "qh", text: "K/T TRƯỞNG BAN", col: 1, bold: true },
      { key: "cv", text: "PHÓ TRƯỞNG BAN", col: 1 },
      { key: "ten", text: "Nguyễn Văn A", col: 1, bold: true },
      { key: "xn", text: "Xác nhận chữ ký của đồng chí Trần Văn B", col: 1 },
    ].map(({ key, ...p }) => ({ row: 0, ...p, ...over[key] }));

  it("đúng thì không lỗi (kể cả ký hiệu người soạn và dòng xác nhận)", () => {
    expect(run(doc(footer()))).toEqual([]);
  });

  it("\"Nơi nhận\" phải gạch chân, không đậm, không nghiêng, cỡ 14", () => {
    for (const patch of [{ underline: false }, { bold: true }, { italic: true }, { sizePt: 13 }]) {
      const d = doc(footer({ nn: patch }));
      expect(ids(run(d))).toEqual(["footer.noiNhan.format"]);
      expectFixWorks(d, "footer.noiNhan.format");
    }
  });

  it("danh sách nơi nhận: dấu chấm phẩy, dấu chấm cuối, gạch đầu dòng, cỡ chữ", () => {
    expect(ids(run(doc(footer({ a: { text: "- Như trên" } }))))).toEqual(["footer.noiNhan.punctuation"]);
    expect(ids(run(doc(footer({ b: { text: "- Lưu: VT, TCTU;" } }))))).toEqual(["footer.noiNhan.punctuation"]);
    expect(ids(run(doc(footer({ a: { text: "Như trên;" } }))))).toEqual(["footer.noiNhan.dash"]);
    expect(ids(run(doc(footer({ a: { sizePt: 14 } }))))).toEqual(["footer.noiNhan.size"]);
    expectFixWorks(doc(footer({ a: { text: "- Như trên" } })), "footer.noiNhan.punctuation");
    expectFixWorks(doc(footer({ b: { text: "- Lưu: VT, TCTU;" } })), "footer.noiNhan.punctuation");
    expectFixWorks(doc(footer({ a: { sizePt: 14 } })), "footer.noiNhan.size");
  });

  it("chức vụ không được đậm, phải in hoa, cỡ 14", () => {
    expect(ids(run(doc(footer({ cv: { bold: true } }))))).toEqual(["footer.chucVu.bold"]);
    expect(ids(run(doc(footer({ cv: { text: "Phó Trưởng ban" } }))))).toEqual(["footer.chucVu.upper"]);
    expect(ids(run(doc(footer({ cv: { sizePt: 13 } }))))).toEqual(["footer.chucVu.size"]);
    expectFixWorks(doc(footer({ cv: { bold: true } })), "footer.chucVu.bold");
    expectFixWorks(doc(footer({ cv: { text: "Phó Trưởng ban" } })), "footer.chucVu.upper");
    expectFixWorks(doc(footer({ cv: { sizePt: 13 } })), "footer.chucVu.size");
  });

  it("quyền hạn viết kiểu Chính quyền được sửa thành T/M, K/T, T/L", () => {
    for (const [wrong, right] of [["TM. BAN THƯỜNG VỤ", "T/M BAN THƯỜNG VỤ"], ["KT. TRƯỞNG BAN", "K/T TRƯỞNG BAN"], ["TL. TRƯỞNG BAN", "T/L TRƯỞNG BAN"]]) {
      const d = doc(footer({ qh: { text: wrong } }));
      const issue = run(d).find((i) => i.ruleId === "footer.quyenHan.format");
      expect(issue?.fix && applyFix(d, issue.fix).paragraphs[4].resolved.text).toBe(right);
      expectFixWorks(d, "footer.quyenHan.format");
    }
  });

  it("quyền hạn TUQ. không có trong chuẩn Đảng và không tự sửa", () => {
    const issue = run(doc(footer({ qh: { text: "TUQ. TRƯỞNG BAN" } }))).find((i) => i.ruleId === "footer.quyenHan.format");
    expect(issue).toBeDefined();
    expect(issue?.fix).toBeUndefined();
  });

  it("họ tên: phải đậm, không in hoa, không học hàm", () => {
    expect(ids(run(doc(footer({ ten: { bold: false } }))))).toEqual(["footer.hoTen.bold"]);
    expect(ids(run(doc(footer({ ten: { text: "NGUYỄN VĂN A" } }))))).toEqual(["footer.hoTen.case"]);
    expect(ids(run(doc(footer({ ten: { text: "TS. Nguyễn Văn A" } }))))).toEqual(["footer.hoTen.hocHam"]);
    expectFixWorks(doc(footer({ ten: { bold: false } })), "footer.hoTen.bold");
  });
});

describe("trộn dấu hiệu của hai chuẩn", () => {
  const QH = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";

  it("văn bản Đảng có cả quốc hiệu: cảnh báo trộn, không tự sửa", () => {
    const issues = checkMixedStandards(doc([{ text: TITLE, col: 1, bold: true }, { text: QH, col: 0, bold: true }]), HD05);
    expect(ids(issues)).toEqual(["header.mixed"]);
    expect(issues[0].severity).toBe("warning");
    expect(issues[0].fix).toBeUndefined();
  });

  it("văn bản Đảng chỉ có quốc hiệu: báo là tiêu đề của chuẩn khác", () => {
    const issues = checkMixedStandards(doc([{ text: QH, bold: true }]), HD05);
    expect(ids(issues)).toEqual(["header.otherStandard"]);
    expect(issues[0].severity).toBe("error");
  });

  it("văn bản Chính quyền có thêm tiêu đề Đảng: cảnh báo trộn", () => {
    expect(ids(checkMixedStandards(doc([{ text: QH, bold: true }, { text: TITLE, col: 1 }]), ND30))).toEqual(["header.mixed"]);
  });

  it("văn bản đúng chuẩn của mình không bị báo", () => {
    expect(checkMixedStandards(doc([{ text: TITLE, col: 1, bold: true }]), HD05)).toEqual([]);
    expect(checkMixedStandards(doc([{ text: QH, bold: true }]), ND30)).toEqual([]);
  });

  it("cả hai chuẩn đều đăng ký kiểm tra trộn chuẩn", () => {
    for (const profile of [HD05, ND30]) expect(headerFooterRules(profile)).toContain(checkMixedStandards);
  });
});
