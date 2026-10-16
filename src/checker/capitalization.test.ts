import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { checkCapitalization } from "./capitalization";
import { applyFix } from "./fixes";
import { planFixes, simulateFixAll } from "./fixAll";
import type { CheckDocument, CheckParagraph, Issue } from "./types";

function doc(texts: string[], over: Partial<CheckParagraph> = {}): CheckDocument {
  const paragraphs: CheckParagraph[] = texts.map((text, index) => ({
    index,
    inTable: false,
    resolved: { kind: "paragraph", text, align: "both", sizePt: 14 },
    ...over,
  }));
  return { pages: [], paragraphs };
}

const run = (d: CheckDocument, profile = ND30): Issue[] => checkCapitalization(d, profile);
const ids = (issues: Issue[]) => issues.map((i) => i.ruleId);
const fixedText = (text: string, profile = ND30): string => simulateFixAll(doc([text]), (d) => checkCapitalization(d, profile)).doc.paragraphs[0].resolved.text;

describe("viết hoa đầu câu (Phụ lục II, mục I)", () => {
  it("đầu đoạn văn viết thường bị báo và sửa được", () => {
    expect(ids(run(doc(["thực hiện nhiệm vụ được giao trong năm 2026."])))).toEqual(["cap.sentenceStart"]);
    expect(fixedText("thực hiện nhiệm vụ được giao trong năm 2026.")).toBe("Thực hiện nhiệm vụ được giao trong năm 2026.");
  });

  it("bỏ qua nhãn đầu dòng, đánh số và dấu ngoặc khi xét chữ đầu", () => {
    expect(fixedText("- như trên;")).toBe("- Như trên;");
    expect(fixedText("2. tổ chức thực hiện.")).toBe("2. Tổ chức thực hiện.");
    expect(fixedText("a) giao Sở Tài chính chủ trì.")).toBe("a) Giao Sở Tài chính chủ trì.");
    expect(fixedText("“tăng cường kỷ luật.”")).toBe("“Tăng cường kỷ luật.”");
  });

  it("sau dấu chấm, chấm hỏi, chấm than trong đoạn phải viết hoa", () => {
    expect(ids(run(doc(["Đã hoàn thành. tiếp tục triển khai! sao lại thế? được."])))).toEqual(["cap.afterPeriod", "cap.afterPeriod", "cap.afterPeriod"]);
    expect(fixedText("Đã hoàn thành. tiếp tục triển khai! sao lại thế? được.")).toBe("Đã hoàn thành. Tiếp tục triển khai! Sao lại thế? Được.");
  });

  it("không báo nhầm sau chữ viết tắt, số thập phân, dấu chấm lửng", () => {
    expect(run(doc(["Cấp cho TP. Cần Thơ và ThS. Nguyễn Văn A; tỷ lệ 1.5 lần."]))).toEqual([]);
    expect(run(doc(["Gồm nhiều nội dung v.v. và các việc khác."]))).toEqual([]);
    expect(run(doc(["Các việc như: báo cáo, tổng hợp… và theo dõi."]))).toEqual([]);
  });

  it("không xét đoạn trong bảng, đoạn in hoa toàn bộ và đoạn trống", () => {
    expect(run(doc(["thực hiện nhiệm vụ."], { inTable: true, table: { index: 0, row: 0, col: 0 } }))).toEqual([]);
    expect(run(doc(["QUYẾT ĐỊNH:", "", "   "]))).toEqual([]);
  });
});

describe("viết hoa tên cơ quan, tổ chức và chức vụ (Phụ lục II, mục II)", () => {
  it("viết thường hoặc viết hoa thừa tên cố định bị báo, sửa về đúng cách viết", () => {
    expect(fixedText("Báo cáo gửi ủy ban nhân dân thành phố.")).toBe("Báo cáo gửi Ủy ban nhân dân thành phố.");
    expect(fixedText("Theo đề nghị của Ủy Ban Nhân Dân quận.")).toBe("Theo đề nghị của Ủy ban nhân dân quận.");
    expect(fixedText("Thực hiện chỉ đạo của thủ tướng chính phủ.")).toBe("Thực hiện chỉ đạo của Thủ tướng Chính phủ.");
    expect(fixedText("Nghị quyết của ban chấp hành trung ương.")).toBe("Nghị quyết của Ban Chấp hành Trung ương.");
  });

  it("chấp nhận cả cách bỏ dấu Uỷ và Ủy", () => {
    expect(run(doc(["Gửi Uỷ ban nhân dân tỉnh."]))).toEqual([]);
  });

  it("không nhầm từ chung đứng trong từ khác", () => {
    expect(run(doc(["Quốc hội và Chính phủ thống nhất."]))).toEqual([]);
  });
});

describe("viết hoa tên văn bản và phần, chương, điều, khoản, điểm khi viện dẫn", () => {
  it("tên loại văn bản đi với số hoặc tên riêng phải viết hoa chữ đầu", () => {
    expect(fixedText("Căn cứ luật Đất đai và nghị định số 30/2020/NĐ-CP.")).toBe("Căn cứ Luật Đất đai và Nghị định số 30/2020/NĐ-CP.");
    expect(fixedText("Theo bộ luật Dân sự.")).toBe("Theo Bộ luật Dân sự.");
  });

  it("chuẩn Chính quyền: viện dẫn Phần, Chương, Mục, Tiểu mục, Điều viết hoa; khoản, điểm viết thường (issue #1)", () => {
    expect(fixedText("Căn cứ điểm a khoản 1 điều 5 chương II của Luật.")).toBe("Căn cứ điểm a khoản 1 Điều 5 Chương II của Luật.");
    expect(fixedText("Theo tiểu mục 2 mục 5 phần I.")).toBe("Theo Tiểu mục 2 Mục 5 Phần I.");
    expect(run(doc(["Căn cứ điểm a khoản 2 Điều 103 Mục 5 Chương XII Phần I của Bộ luật Hình sự."]))).toEqual([]);
  });

  it("khoản, điểm viết hoa giữa câu khi viện dẫn bị báo, sửa về viết thường; đầu câu thì giữ viết hoa", () => {
    expect(ids(run(doc(["Căn cứ Điểm a Khoản 2 Điều 103 của Luật."])))).toEqual(["cap.referenceLower", "cap.referenceLower"]);
    expect(fixedText("Căn cứ Điểm a Khoản 2 Điều 103 của Luật.")).toBe("Căn cứ điểm a khoản 2 Điều 103 của Luật.");
    expect(run(doc(["Khoản 1 Điều 5 quy định như sau."]))).toEqual([]);
    expect(run(doc(["Đã rõ. Điểm a khoản 1 Điều 5 quy định như sau."]))).toEqual([]);
  });

  it("không nhầm khoản tiền, điểm số", () => {
    expect(run(doc(["Chi một khoản 2 triệu đồng; đạt điểm 9 trên 10."]))).toEqual([]);
  });

  it("chuẩn Đảng không áp quy tắc viện dẫn điều, khoản của Nghị định 30", () => {
    expect(ids(run(doc(["Căn cứ điểm a khoản 1 điều 5."]), HD05))).toEqual([]);
  });
});

describe("cách sửa nhiều chỗ trên một đoạn", () => {
  it("các lỗi viết hoa trên cùng một đoạn gộp thành một lần sửa", () => {
    const issues = run(doc(["thực hiện. tiếp tục theo luật Đất đai."]));
    expect(issues.length).toBe(3);
    expect(planFixes(issues).fixes).toHaveLength(1);
    const fixed = applyFix(doc(["thực hiện. tiếp tục theo luật Đất đai."]), planFixes(issues).fixes[0]);
    expect(fixed.paragraphs[0].resolved.text).toBe("Thực hiện. Tiếp tục theo Luật Đất đai.");
  });

  it("mọi lỗi viết hoa là cảnh báo, có căn cứ theo chuẩn", () => {
    for (const issue of run(doc(["thực hiện. tiếp tục."]))) {
      expect(issue.severity).toBe("warning");
      expect(issue.reference).toContain("Phụ lục II");
    }
  });
});

describe("đoạn có quy cách riêng", () => {
  it("trích yếu căn giữa (văn bản Đảng viết thường) không bị báo đầu câu", () => {
    const d: CheckDocument = { pages: [], paragraphs: [{ index: 0, inTable: false, resolved: { kind: "paragraph", text: "thực hiện quy chế làm việc", align: "center" } }] };
    expect(checkCapitalization(d, HD05)).toEqual([]);
  });
});
