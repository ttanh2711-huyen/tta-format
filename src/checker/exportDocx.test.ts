import AdmZip from "adm-zip";
import { buildReportDocx, crc32, toBase64 } from "./exportDocx";
import type { Issue } from "./types";

const issue = (over: Partial<Issue> = {}): Issue => ({ ruleId: "r", severity: "error", message: "Sai <phông> & cỡ", reference: "Điều 1", location: { kind: "paragraph", index: 2, excerpt: "Đoạn a" }, ...over });
const header = { standardName: "Chính quyền", checkedAt: new Date(2026, 9, 13, 8, 30) };

describe("crc32", () => {
  it("khớp giá trị chuẩn của chuỗi 123456789", () => {
    expect(crc32(new TextEncoder().encode("123456789"))).toBe(0xcbf43926);
  });
});

describe("báo cáo .docx", () => {
  const zip = (issues: Issue[]) => new AdmZip(Buffer.from(buildReportDocx(issues, header)));

  it("là tệp zip hợp lệ có đủ các phần của một tài liệu Word", () => {
    const names = zip([issue()]).getEntries().map((e) => e.entryName);
    expect(names).toEqual(expect.arrayContaining(["[Content_Types].xml", "_rels/.rels", "word/document.xml"]));
  });

  it("nội dung có tiêu đề, từng lỗi, vị trí, căn cứ; ký tự đặc biệt của XML được thoát", () => {
    const xml = zip([issue({ fix: { kind: "font", index: 2, bold: true, label: "Đặt đậm" } })]).readAsText("word/document.xml");
    expect(xml).toContain("BÁO CÁO KIỂM TRA VĂN BẢN THEO CHUẨN CHÍNH QUYỀN");
    expect(xml).toContain("Sai &lt;phông&gt; &amp; cỡ");
    expect(xml).toContain("Điều 1");
    expect(xml).toContain("Đặt đậm");
    expect(xml).not.toContain("<phông>");
  });

  it("không có lỗi thì ghi rõ", () => {
    expect(zip([]).readAsText("word/document.xml")).toContain("Không thấy lỗi.");
  });

  it("base64 giải mã lại đúng tệp zip", () => {
    const bytes = buildReportDocx([issue()], header);
    expect(Buffer.from(toBase64(bytes), "base64").equals(Buffer.from(bytes))).toBe(true);
  });
});
