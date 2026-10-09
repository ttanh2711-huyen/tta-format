import { readFileSync } from "node:fs";
import { join } from "node:path";
import { accessibleNameFor } from "./a11y";

describe("tên điều khiển cho trình đọc màn hình", () => {
  it("điều khiển đã có nhãn hoặc aria thì không gắn thêm", () => {
    expect(accessibleNameFor({ labelCount: 1, placeholder: "Gợi ý" })).toBeNull();
    expect(accessibleNameFor({ labelCount: 0, ariaLabel: "Đã có", placeholder: "Gợi ý" })).toBeNull();
    expect(accessibleNameFor({ labelCount: 0, ariaLabelledBy: "id-khac", placeholder: "Gợi ý" })).toBeNull();
  });

  it("chưa có nhãn thì lấy chữ gợi ý, rồi tiêu đề", () => {
    expect(accessibleNameFor({ labelCount: 0, placeholder: "  Tên cơ quan  " })).toBe("Tên cơ quan");
    expect(accessibleNameFor({ labelCount: 0, title: "Chọn dấu" })).toBe("Chọn dấu");
    expect(accessibleNameFor({ labelCount: 0, placeholder: "A", title: "B" })).toBe("A");
  });

  it("không có gì để lấy hoặc ô ẩn thì không gắn", () => {
    expect(accessibleNameFor({ labelCount: 0 })).toBeNull();
    expect(accessibleNameFor({ labelCount: 0, placeholder: "  " })).toBeNull();
    expect(accessibleNameFor({ labelCount: 0, type: "hidden", placeholder: "x" })).toBeNull();
  });
});

describe("tương phản màu của giao diện (WCAG AA 4,5:1 cho chữ thường)", () => {
  const html = readFileSync(join(__dirname, "..", "taskpane", "taskpane.html"), "utf8");

  function luminance(hex: string): number {
    const h = hex.replace("#", "");
    const full = h.length === 3 ? [...h].map((c) => c + c).join("") : h;
    const [r, g, b] = [0, 2, 4].map((i) => parseInt(full.slice(i, i + 2), 16) / 255).map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
  }
  function ratio(a: string, b: string): number {
    const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x);
    return (hi + 0.05) / (lo + 0.05);
  }

  // Chữ trắng chỉ đặt trên nền đậm của nút; còn lại là chữ tối trên nền trắng hoặc nền nhạt.
  const textColors = [...new Set([...html.matchAll(/(?<![-\w])color:\s*(#[0-9a-fA-F]{3,6})/g)].map((m) => m[1].toLowerCase()))].filter((c) => c !== "#fff" && c !== "#ffffff");
  const lightBackgrounds = [...new Set([...html.matchAll(/background:\s*(#[0-9a-fA-F]{3,6})/g)].map((m) => m[1].toLowerCase()))].filter((c) => luminance(c) > 0.5);

  it("có màu chữ để kiểm tra", () => {
    expect(textColors.length).toBeGreaterThan(3);
    expect(lightBackgrounds.length).toBeGreaterThan(1);
  });

  it.each(textColors)("chữ %s đủ tương phản trên nền trắng và mọi nền nhạt", (color) => {
    expect(ratio(color, "#ffffff")).toBeGreaterThanOrEqual(4.5);
    for (const background of lightBackgrounds) expect(ratio(color, background)).toBeGreaterThanOrEqual(4.5);
  });

  it("chữ trắng trên nút nền đậm đủ tương phản", () => {
    for (const background of [...html.matchAll(/background:\s*(#[0-9a-fA-F]{3,6});\s*color:\s*#fff/g)].map((m) => m[1])) expect(ratio("#ffffff", background)).toBeGreaterThanOrEqual(4.5);
  });

  it("có kiểu hiển thị khi điều khiển được chọn bằng bàn phím", () => {
    expect(html).toContain(":focus-visible");
  });

  it("trang khai báo ngôn ngữ tiếng Việt", () => {
    expect(html).toContain('<html lang="vi">');
  });
});
