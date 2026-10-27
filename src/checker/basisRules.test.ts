import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { checkBasisBlocks, findBasisBlocks } from "./basisRules";
import { applyFix } from "./fixes";
import type { CheckDocument, CheckParagraph, Issue } from "./types";

interface P {
  text: string;
  align?: "center" | "both" | "left";
  italic?: boolean;
  bold?: boolean;
  table?: boolean;
}

function doc(items: P[]): CheckDocument {
  const paragraphs: CheckParagraph[] = items.map((p, index) => ({
    index,
    inTable: p.table === true,
    table: p.table ? { index: 0, row: 0, col: 0 } : undefined,
    resolved: { kind: "paragraph", text: p.text, align: p.align ?? "both", italic: p.italic, bold: p.bold, sizePt: 14 },
  }));
  return { pages: [], paragraphs };
}

const ids = (issues: Issue[]) => issues.map((i) => i.ruleId);
const header: P[] = [{ text: "ỦY BAN NHÂN DÂN XÃ A", table: true }, { text: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", table: true }];

describe("tìm khối căn cứ", () => {
  it("ngay dưới tên loại, trích yếu của văn bản có tên loại (kế hoạch, hướng dẫn…); dừng ở đoạn thân đầu tiên", () => {
    const d = doc([
      ...header,
      { text: "KẾ HOẠCH", align: "center", bold: true },
      { text: "Triển khai công tác chuyển đổi số năm 2026", align: "center", bold: true },
      { text: "Căn cứ Nghị quyết số 57-NQ/TW;", italic: true },
      { text: "Căn cứ Kế hoạch số 12/KH-UBND.", italic: true },
      { text: "Ủy ban nhân dân xã ban hành Kế hoạch như sau:" },
    ]);
    expect(findBasisBlocks(d).map((block) => block.map((p) => p.index))).toEqual([[4, 5]]);
  });

  it("công văn (không có tên loại) không có phần căn cứ: câu mở đầu bằng 'Căn cứ' là thân văn bản", () => {
    expect(findBasisBlocks(doc([...header, { text: "Căn cứ Kế hoạch số 12/KH-UBND, Phòng báo cáo như sau:" }]))).toEqual([]);
  });

  it("câu dẫn kết thúc bằng dấu hai chấm không thuộc khối căn cứ", () => {
    const d = doc([...header, { text: "BÁO CÁO", align: "center" }, { text: "Căn cứ Luật A;", italic: true }, { text: "Căn cứ tình hình thực tế, đơn vị báo cáo như sau:" }]);
    expect(findBasisBlocks(d).map((block) => block.map((p) => p.index))).toEqual([[3]]);
  });

  it("khối trên dòng chốt và khối dưới tên loại trùng nhau thì chỉ tính một lần", () => {
    const d = doc([...header, { text: "QUYẾT ĐỊNH", align: "center" }, { text: "Về việc A", align: "center" }, { text: "CHỦ TỊCH ỦY BAN NHÂN DÂN XÃ", align: "center" }, { text: "Căn cứ Luật A;", italic: true }, { text: "Theo đề nghị của Văn phòng.", italic: true }, { text: "QUYẾT ĐỊNH:", align: "center", bold: true }]);
    expect(findBasisBlocks(d)).toHaveLength(1);
  });
});

describe("căn cứ theo Nghị định 30 (Phụ lục I, II.6)", () => {
  const base: P[] = [...header, { text: "KẾ HOẠCH", align: "center" }, { text: "Căn cứ Luật A;", italic: true }, { text: "Căn cứ Nghị định B.", italic: true }, { text: "Nội dung." }];

  it("đúng: in nghiêng, không gạch đầu dòng, ';' và '.' cuối", () => {
    expect(checkBasisBlocks(doc(base), ND30)).toEqual([]);
  });

  it("gạch đầu dòng, chữ đứng, căn cứ cuối thiếu dấu chấm: báo và sửa được", () => {
    const d = doc([...header, { text: "KẾ HOẠCH", align: "center" }, { text: "- Căn cứ Luật A;" }, { text: "Căn cứ Nghị định B;", italic: true }, { text: "Nội dung." }]);
    expect(ids(checkBasisBlocks(d, ND30)).sort()).toEqual(["basis.dash", "basis.italic", "basis.punctuation"]);
    let fixed = d;
    for (const issue of checkBasisBlocks(d, ND30)) fixed = applyFix(fixed, issue.fix!);
    expect(checkBasisBlocks(fixed, ND30)).toEqual([]);
  });
});

describe("căn cứ theo Hướng dẫn 05 (6.2.3, 6.2.2)", () => {
  const party = (lines: P[]): CheckDocument => doc([{ text: "ĐẢNG CỘNG SẢN VIỆT NAM", table: true }, { text: "NGHỊ QUYẾT", align: "center" }, { text: "về công tác A", align: "center" }, { text: "-----", align: "center" }, ...lines, { text: "Nội dung." }]);

  it("đúng: gạch ngang đầu mỗi căn cứ, chữ đứng, ';' và căn cứ cuối '.'", () => {
    expect(checkBasisBlocks(party([{ text: "- Căn cứ Điều lệ Đảng;" }, { text: "- Căn cứ Quy định số 399-QĐ/TW." }]), HD05)).toEqual([]);
  });

  it("thiếu gạch ngang, in nghiêng, căn cứ cuối sai dấu: báo và sửa được", () => {
    const d = party([{ text: "Căn cứ Điều lệ Đảng;", italic: true }, { text: "- Căn cứ Quy định số 399-QĐ/TW;" }]);
    expect(ids(checkBasisBlocks(d, HD05)).sort()).toEqual(["basis.dash", "basis.punctuation", "basis.upright"]);
    let fixed = d;
    for (const issue of checkBasisBlocks(d, HD05)) fixed = applyFix(fixed, issue.fix!);
    expect(checkBasisBlocks(fixed, HD05)).toEqual([]);
  });
});
