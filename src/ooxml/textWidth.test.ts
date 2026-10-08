import { textWidthMm } from "./textWidth";

describe("textWidthMm: ước lượng độ rộng chữ Times New Roman", () => {
  it("chuỗi rỗng rộng 0", () => {
    expect(textWidthMm("", 13, false)).toBe(0);
  });

  it("tỉ lệ với cỡ chữ", () => {
    expect(textWidthMm("ABC", 26, false)).toBeCloseTo(2 * textWidthMm("ABC", 13, false), 5);
  });

  it("chữ đậm rộng hơn chữ thường", () => {
    expect(textWidthMm("CỘNG HÒA", 13, true)).toBeGreaterThan(textWidthMm("CỘNG HÒA", 13, false));
  });

  it("dấu tiếng Việt không đổi độ rộng của chữ cái gốc", () => {
    expect(textWidthMm("ĐƯỜNG", 13, true)).toBeCloseTo(textWidthMm("DUONG", 13, true), 5);
    expect(textWidthMm("Độc lập", 14, true)).toBeCloseTo(textWidthMm("Doc lap", 14, true), 5);
  });

  it("khớp độ dài đường kẻ trong file mẫu (tiêu ngữ 14 pt đậm: 59 mm) trong ±6%", () => {
    const width = textWidthMm("Độc lập - Tự do - Hạnh phúc", 14, true);
    expect(width).toBeGreaterThan(59 * 0.94);
    expect(width).toBeLessThan(59 * 1.06);
  });

  it("tên cơ quan 13 pt đậm: đường kẻ 17,5 mm của mẫu nằm trong khoảng 1/3 đến 1/2 độ rộng", () => {
    const width = textWidthMm("CHI CỤC THỦY LỢI", 13, true);
    expect(17.5 / width).toBeGreaterThanOrEqual(1 / 3 - 0.02);
    expect(17.5 / width).toBeLessThanOrEqual(1 / 2 + 0.02);
  });
});
