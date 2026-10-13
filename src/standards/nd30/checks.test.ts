import { applyFix, fixIndex, isFormatOnly } from "../../checker/fixes";
import { runChecks } from "../../checker/engine";
import type { CheckDocument, CheckParagraph, Issue } from "../../checker/types";
import { buildCheckDocument } from "../../checker/documentModel";
import { assembleDocumentBody, BODY_MARKER } from "../../ooxml/documentAssembler";
import { buildStyleSpecs } from "../../ooxml/styleSpecs";
import { parseStyleTable } from "../../ooxml/styleResolver";
import { toStyleDefinitionsXml } from "../../ooxml/styleXml";
import { extractDocModel } from "../../rules/docModel";
import { ND30 } from "./profile";
import { checkNd30HeaderFooter } from "./checks";

interface P {
  text: string;
  col?: number;
  row?: number;
  bold?: boolean;
  italic?: boolean;
  sizePt?: number;
}

/** Tài liệu một bảng: mỗi đoạn có ô (hàng, cột) và định dạng; đoạn lần lượt đánh số từ 0. */
function doc(paragraphs: P[]): CheckDocument {
  const list: CheckParagraph[] = paragraphs.map((p, index) => ({
    index,
    inTable: true,
    table: { index: 0, row: p.row ?? 0, col: p.col ?? 0 },
    resolved: { kind: "paragraph", text: p.text, bold: p.bold, italic: p.italic, sizePt: p.sizePt ?? 13 },
  }));
  return { pages: [], paragraphs: list };
}

const run = (d: CheckDocument): Issue[] => checkNd30HeaderFooter(d, ND30);
const ids = (issues: Issue[]) => issues.map((i) => i.ruleId);
const QH = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
const TN = "Độc lập - Tự do - Hạnh phúc";

/** Kiểm tra: áp cách sửa thì lỗi biến mất và không đoạn nào khác thay đổi. */
function expectFixWorks(d: CheckDocument, ruleId: string): void {
  const issue = run(d).find((i) => i.ruleId === ruleId);
  expect(issue?.fix).toBeDefined();
  const fixed = applyFix(d, issue!.fix!);
  expect(run(fixed).filter((i) => i.ruleId === ruleId && i.location.kind === "paragraph" && i.location.index === fixIndex(issue!.fix!))).toEqual([]);
  fixed.paragraphs.forEach((p, i) => {
    if (i !== fixIndex(issue!.fix!)) expect(p).toBe(d.paragraphs[i]);
  });
}

describe("quốc hiệu và tiêu ngữ", () => {
  it("đúng chữ, đậm, đúng cỡ thì không có lỗi", () => {
    expect(run(doc([{ text: QH, bold: true }, { text: TN, bold: true, sizePt: 14 }]))).toEqual([]);
  });

  it.each([
    ["sai chữ cuối", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAN"],
    ["mất dấu", "CONG HOA XA HOI CHU NGHIA VIET NAM"],
    ["viết thường", "Cộng hòa xã hội chủ nghĩa Việt Nam"],
    ["dấu thanh kiểu cũ", "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM"],
  ])("quốc hiệu %s bị báo kèm cách sửa", (_name, text) => {
    const d = doc([{ text, bold: true }]);
    expect(ids(run(d))).toEqual(["header.quocHieu.text"]);
    expect(run(d)[0].expected).toBe(QH);
    expectFixWorks(d, "header.quocHieu.text");
  });

  it("quốc hiệu không đậm hoặc cỡ chữ ngoài khoảng", () => {
    expect(ids(run(doc([{ text: QH }])))).toEqual(["header.quocHieu.bold"]);
    expect(ids(run(doc([{ text: QH, bold: true, sizePt: 16 }])))).toEqual(["header.quocHieu.size"]);
    expectFixWorks(doc([{ text: QH }]), "header.quocHieu.bold");
    expectFixWorks(doc([{ text: QH, bold: true, sizePt: 16 }]), "header.quocHieu.size");
  });

  it("tiêu ngữ dùng dấu gạch dài hoặc viết thiếu dấu bị báo", () => {
    for (const text of ["Độc lập – Tự do – Hạnh phúc", "Độc lập - Tự do - Hạnh phuc", "Độc lập - Tự do - Hạnh phúc."]) {
      const d = doc([{ text, bold: true, sizePt: 14 }]);
      expect(ids(run(d))).toEqual(["header.tieuNgu.text"]);
      expectFixWorks(d, "header.tieuNgu.text");
    }
  });

  it("tiêu ngữ không đậm", () => {
    expect(ids(run(doc([{ text: TN, sizePt: 14 }])))).toEqual(["header.tieuNgu.bold"]);
  });

  it("đoạn không liên quan trong bảng không bị báo nhầm", () => {
    expect(run(doc([{ text: "Cộng hòa" }, { text: "Độc lập" }, { text: "SỞ NÔNG NGHIỆP VÀ MÔI TRƯỜNG", bold: true }]))).toEqual([]);
  });
});

describe("số và ký hiệu", () => {
  it.each(["Số: 05/QĐ-UBND", "Số: 127/QĐ-UBND", "Số: 15/BNV-VP", "Số:        /CCTL-ĐĐPCTT"])("%s đúng", (text) => {
    expect(run(doc([{ text }]))).toEqual([]);
  });

  it.each([
    ["số nhỏ hơn 10 thiếu số 0", "Số: 5/QĐ-UBND", "Số: 05/QĐ-UBND"],
    ["thiếu dấu hai chấm", "Số 05/QĐ-UBND", "Số: 05/QĐ-UBND"],
    ["khoảng trắng quanh dấu /", "Số: 05 / QĐ-UBND", "Số: 05/QĐ-UBND"],
    ["khoảng trắng quanh dấu -", "Số: 05/QĐ - UBND", "Số: 05/QĐ-UBND"],
    ["thiếu khoảng trắng sau Số:", "Số:05/QĐ-UBND", "Số: 05/QĐ-UBND"],
  ])("%s", (_name, text, expected) => {
    const d = doc([{ text }]);
    const issues = run(d);
    expect(ids(issues)).toEqual(["header.so.format"]);
    expect(issues[0].expected).toBe(expected);
    expectFixWorks(d, "header.so.format");
  });

  it("nhiều lỗi cùng lúc được nêu đủ trong một thông báo", () => {
    const [issue] = run(doc([{ text: "Số 5 / QĐ - UBND" }]));
    expect(issue.message).toContain("dấu hai chấm");
    expect(issue.message).toContain("số 0");
    expect(issue.message).toContain("\"/\"");
    expect(issue.expected).toBe("Số: 05/QĐ-UBND");
  });

  it("dòng bắt đầu bằng chữ Số nhưng không phải số và ký hiệu thì bỏ qua", () => {
    expect(run(doc([{ text: "Số lượng người tham dự" }, { text: "Số tiền 5/10 đồng" }]))).toEqual([]);
  });
});

describe("ngày tháng", () => {
  it.each(["Hà Nội, ngày 05 tháng 01 năm 2020", "Cần Thơ, ngày 15 tháng 12 năm 2026", "Cần Thơ, ngày 05 tháng 3 năm 2026", "Cần Thơ, ngày      tháng 10 năm 2026", "Cần Thơ, ngày      tháng      năm      "])("%s đúng", (text) => {
    expect(run(doc([{ text }]))).toEqual([]);
  });

  it.each([
    ["ngày thiếu số 0", "Hà Nội, ngày 5 tháng 01 năm 2020", "Hà Nội, ngày 05 tháng 01 năm 2020"],
    ["tháng 1 thiếu số 0", "Hà Nội, ngày 05 tháng 1 năm 2020", "Hà Nội, ngày 05 tháng 01 năm 2020"],
    ["tháng 2 thiếu số 0", "Hà Nội, ngày 10 tháng 2 năm 2020", "Hà Nội, ngày 10 tháng 02 năm 2020"],
    ["tháng từ 3 trở đi thừa số 0", "Hà Nội, ngày 05 tháng 03 năm 2020", "Hà Nội, ngày 05 tháng 3 năm 2020"],
  ])("%s", (_name, text, expected) => {
    const d = doc([{ text }]);
    const issues = run(d);
    expect(ids(issues)).toEqual(["header.date.format"]);
    expect(issues[0].expected).toBe(expected);
    expectFixWorks(d, "header.date.format");
  });

  it("ngày không tồn tại bị báo, không có cách sửa tự động", () => {
    const issues = run(doc([{ text: "Hà Nội, ngày 31 tháng 02 năm 2020" }]));
    expect(ids(issues)).toEqual(["header.date.invalid"]);
    expect(issues[0].fix).toBeUndefined();
    expect(ids(run(doc([{ text: "Hà Nội, ngày 31 tháng 04 năm 2020" }])))).toContain("header.date.invalid");
  });

  it("năm không đủ 4 chữ số là cảnh báo", () => {
    const issues = run(doc([{ text: "Hà Nội, ngày 05 tháng 01 năm 20" }]));
    expect(ids(issues)).toEqual(["header.date.year"]);
    expect(issues[0].severity).toBe("warning");
  });

  it("câu có chữ ngày tháng năm trong đoạn thân (ngoài bảng) không bị báo", () => {
    const outside: CheckDocument = { pages: [], paragraphs: [{ index: 0, inTable: false, resolved: { kind: "paragraph", text: "Hà Nội, ngày 5 tháng 1 năm 2020" } }] };
    expect(run(outside)).toEqual([]);
  });
});

describe("nơi nhận và chữ ký", () => {
  const clean = (): P[] => [
    { text: "Nơi nhận:", bold: true, italic: true, sizePt: 12 },
    { text: "- Như trên;", sizePt: 11 },
    { text: "- Lưu: VT.", sizePt: 11 },
    { text: "KT. CHI CỤC TRƯỞNG", col: 1, bold: true, sizePt: 14 },
    { text: "PHÓ CHI CỤC TRƯỞNG", col: 1, bold: true, sizePt: 14 },
    { text: "Mai Hiếu Hiền", col: 1, bold: true, sizePt: 14 },
  ];
  const with_ = (patch: (list: P[]) => void): CheckDocument => {
    const list = clean();
    patch(list);
    return doc(list);
  };

  it("khối cuối đúng thể thức không có lỗi", () => {
    expect(run(doc(clean()))).toEqual([]);
  });

  it("'Nơi nhận' thiếu đậm hoặc nghiêng hoặc sai cỡ", () => {
    const d = with_((l) => (l[0] = { text: "Nơi nhận:", sizePt: 14 }));
    expect(ids(run(d))).toEqual(["footer.noiNhan.format"]);
    expectFixWorks(d, "footer.noiNhan.format");
  });

  it("mục nơi nhận sai dấu cuối: thiếu ';' ở mục giữa, thiếu '.' ở mục cuối", () => {
    const d = with_((l) => {
      l[1] = { text: "- Như trên", sizePt: 11 };
      l[2] = { text: "- Lưu: VT", sizePt: 11 };
    });
    expect(ids(run(d))).toEqual(["footer.noiNhan.punctuation", "footer.noiNhan.punctuation"]);
    expect(run(d).map((i) => i.expected)).toEqual([";", "."]);
    expectFixWorks(d, "footer.noiNhan.punctuation");
  });

  it("mục cuối kết thúc bằng ';' thay vì '.' và ngược lại", () => {
    const d = with_((l) => {
      l[1] = { text: "- Như trên.", sizePt: 11 };
      l[2] = { text: "- Lưu: VT;", sizePt: 11 };
    });
    expect(ids(run(d))).toEqual(["footer.noiNhan.punctuation", "footer.noiNhan.punctuation"]);
  });

  it("mục nơi nhận không có gạch đầu dòng là cảnh báo", () => {
    const d = with_((l) => (l[1] = { text: "Như trên;", sizePt: 11 }));
    const issues = run(d);
    expect(ids(issues)).toEqual(["footer.noiNhan.dash"]);
    expect(issues[0].severity).toBe("warning");
  });

  it("danh sách nơi nhận sai cỡ chữ", () => {
    const d = with_((l) => (l[1] = { text: "- Như trên;", sizePt: 13 }));
    expect(ids(run(d))).toEqual(["footer.noiNhan.size"]);
    expectFixWorks(d, "footer.noiNhan.size");
  });

  it("chức vụ viết thường bị báo và sửa được bằng in hoa", () => {
    const d = with_((l) => (l[4] = { text: "Phó Chi cục trưởng", col: 1, bold: true, sizePt: 14 }));
    expect(ids(run(d))).toEqual(["footer.chucVu.upper"]);
    expect(run(d)[0].fix).toMatchObject({ kind: "replace", replacement: "PHÓ CHI CỤC TRƯỞNG" });
    expectFixWorks(d, "footer.chucVu.upper");
  });

  it("chức vụ không đậm hoặc sai cỡ", () => {
    const notBold = with_((l) => (l[3] = { text: "KT. CHI CỤC TRƯỞNG", col: 1, sizePt: 14 }));
    expect(ids(run(notBold))).toEqual(["footer.chucVu.bold"]);
    expectFixWorks(notBold, "footer.chucVu.bold");
    const bigger = with_((l) => (l[3] = { text: "KT. CHI CỤC TRƯỞNG", col: 1, bold: true, sizePt: 16 }));
    expect(ids(run(bigger))).toEqual(["footer.chucVu.size"]);
  });

  it.each(["KT CHI CỤC TRƯỞNG", "TM ỦY BAN NHÂN DÂN", "TL GIÁM ĐỐC", "TUQ GIÁM ĐỐC", "Q CHI CỤC TRƯỞNG"])("quyền hạn thiếu dấu chấm: %s", (text) => {
    const d = with_((l) => (l[3] = { text, col: 1, bold: true, sizePt: 14 }));
    expect(ids(run(d))).toEqual(["footer.quyenHan.format"]);
    expectFixWorks(d, "footer.quyenHan.format");
  });

  it("quyền hạn đúng dấu chấm và chức vụ bắt đầu bằng chữ Q... không bị báo nhầm", () => {
    expect(run(with_((l) => (l[3] = { text: "QUYỀN CHI CỤC TRƯỞNG", col: 1, bold: true, sizePt: 14 })))).toEqual([]);
    expect(run(with_((l) => (l[3] = { text: "Q. CHI CỤC TRƯỞNG", col: 1, bold: true, sizePt: 14 })))).toEqual([]);
  });

  it("họ tên có học hàm hoặc học vị bị báo và sửa được bằng cách bỏ học hàm", () => {
    const issues = run(with_((l) => (l[5] = { text: "TS. Mai Hiếu Hiền", col: 1, bold: true, sizePt: 14 })));
    expect(ids(issues)).toEqual(["footer.hoTen.hocHam"]);
    expect(issues[0].fix).toMatchObject({ kind: "replace", replacement: "Mai Hiếu Hiền" });
  });

  it("họ tên in hoa toàn bộ bị báo; họ tên không đậm bị báo và sửa được", () => {
    expect(ids(run(with_((l) => (l[5] = { text: "MAI HIẾU HIỀN", col: 1, bold: true, sizePt: 14 }))))).toEqual(["footer.hoTen.case"]);
    const d = with_((l) => (l[5] = { text: "Mai Hiếu Hiền", col: 1, sizePt: 14 }));
    expect(ids(run(d))).toEqual(["footer.hoTen.bold"]);
    expectFixWorks(d, "footer.hoTen.bold");
  });

  it("chỉ có ô nơi nhận (không có ô chữ ký) thì chỉ kiểm tra nơi nhận", () => {
    expect(run(doc(clean().slice(0, 3)))).toEqual([]);
  });

  it("không có chữ 'Nơi nhận' thì không báo gì về khối cuối", () => {
    expect(run(doc([{ text: "KT chi cục trưởng", col: 1 }, { text: "Ai Đó", col: 1 }]))).toEqual([]);
  });
});

describe("cách sửa", () => {
  it("chỉ sửa định dạng (font) mới được chạy tự động trong Chỉnh toàn bộ; sửa chữ thì không", () => {
    expect(isFormatOnly({ kind: "font", index: 0, bold: true, label: "x" })).toBe(true);
    expect(isFormatOnly({ kind: "replace", index: 0, find: "a", replacement: "b", label: "x" })).toBe(false);
  });

  it("mọi lỗi sửa được đều chỉ trỏ vào đúng đoạn có lỗi", () => {
    const d = doc([{ text: QH }, { text: "Số: 5/QĐ", col: 1 }, { text: "Hà Nội, ngày 5 tháng 1 năm 2020", col: 1, row: 1 }]);
    for (const issue of run(d)) {
      if (issue.fix && issue.location.kind === "paragraph") expect(fixIndex(issue.fix)).toBe(issue.location.index);
    }
  });
});

describe("văn bản dựng từ trình tạo không bị báo sai", () => {
  function docFrom(body: string): CheckDocument {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body.replace(BODY_MARKER, "")}</w:body></w:document>`;
    const styles = parseStyleTable(`<w:styles>${toStyleDefinitionsXml(buildStyleSpecs(ND30))}</w:styles>`);
    return buildCheckDocument(extractDocModel({ documentXml: xml }), styles, []);
  }
  const sizes = { parent: 13, issuer: 13, quocHieu: 13, tieuNgu: 14, so: 13, placeDate: 13, subject: 13 };
  const letter: Parameters<typeof assembleDocumentBody>[0] = {
    header: { parentLines: ["Sở A"], issuer: "Chi cục B", sizes, agencyLineRatio: 0.5, textWidthMm: 160, numberRow: { so: "Số: 05/CCB", placeDate: "Cần Thơ, ngày 05 tháng 10 năm 2026" }, subjectRow: { text: "V/v cử công chức" }, recipientsRow: { label: "Kính gửi:", lines: ["- Sở C;", "- Sở D."], sizePt: 14 } },
    footer: { recipients: ["Như trên", "Sở C"], luu: "VT", authority: "KT.", positionLines: ["Chi cục trưởng", "Phó Chi cục trưởng"], signerName: "Mai Hiếu Hiền", signGapLines: 6, sizes: { noiNhanTieuDe: 12, noiNhanDanhSach: 11, quyenHanChucVu: 14, hoTen: 14 }, textWidthMm: 160 },
  };

  it("công văn dựng từ trình tạo: 0 lỗi", () => {
    expect(runChecks(docFrom(assembleDocumentBody(letter)), ND30, [checkNd30HeaderFooter])).toEqual([]);
  });

  it("khối đầu để trống số và ngày chờ vào sổ: 0 lỗi", () => {
    const blank = { ...letter, header: { ...letter.header, numberRow: { so: "Số:      /CCB", placeDate: "Cần Thơ, ngày      tháng      năm      " } } };
    expect(runChecks(docFrom(assembleDocumentBody(blank)), ND30, [checkNd30HeaderFooter])).toEqual([]);
  });

  it("ký thay, quyền hạn khác: 0 lỗi", () => {
    for (const authority of ["TM.", "Q.", "KT.", "TL.", "TUQ.", ""]) {
      const d = docFrom(assembleDocumentBody({ ...letter, footer: { ...letter.footer, authority } }));
      expect(runChecks(d, ND30, [checkNd30HeaderFooter])).toEqual([]);
    }
  });
});
