import { SAMPLE_TEXT, insertSampleParagraph } from "./insertSample";

describe("insertSampleParagraph", () => {
  it("chèn đúng một đoạn văn bản mẫu vào cuối nội dung", () => {
    const calls: Array<[string, string]> = [];
    const body = {
      insertParagraph: (text: string, where: string) => {
        calls.push([text, where]);
        return {};
      },
    };

    insertSampleParagraph(body);

    expect(calls).toEqual([[SAMPLE_TEXT, "End"]]);
  });

  it("giữ nguyên dấu tiếng Việt ở dạng chuẩn hoá NFC", () => {
    expect(SAMPLE_TEXT).toBe(SAMPLE_TEXT.normalize("NFC"));
  });

  it("chứa các chữ có dấu cần kiểm tra khi thử trong Word", () => {
    expect(SAMPLE_TEXT).toContain("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
    expect(SAMPLE_TEXT).toContain("Độc lập - Tự do - Hạnh phúc");
  });
});
