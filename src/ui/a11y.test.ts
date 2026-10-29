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
  const css = readFileSync(join(__dirname, "..", "taskpane", "taskpane.css"), "utf8");
  const html = readFileSync(join(__dirname, "..", "taskpane", "taskpane.html"), "utf8");
  /** Biến màu khai báo trong khối CSS có bộ chọn `selector` (đúng từng ký tự, ngay trước "{"). */
  function block(selector: string): Map<string, string> {
    const start = css.indexOf(`${selector} {`);
    if (start < 0) throw new Error(`Thiếu khối ${selector} trong taskpane.css`);
    const body = css.slice(start, css.indexOf("}", start));
    return new Map([...body.matchAll(/--([a-z0-9-]+):\s*(#[0-9a-fA-F]{6})\s*;/g)].map((m) => [m[1], m[2].toLowerCase()]));
  }
  // Bảng màu sáng (gốc) và tối (gốc ghi đè bằng khối tối); dải đầu trang thêm biến thể theo chuẩn Đảng.
  const light = block(":root");
  const dark = new Map([...light, ...block(':root[data-theme="dark"]')]);
  const palettes = [
    ["sáng", light],
    ["tối", dark],
  ] as const;
  const headers = [
    ["sáng, Chính quyền", light],
    ["sáng, Đảng", new Map([...light, ...block(':root[data-standard="hd05"]')])],
    ["tối, Chính quyền", dark],
    ["tối, Đảng", new Map([...dark, ...block(':root[data-theme="dark"][data-standard="hd05"]')])],
  ] as const;
  const tokens = light;

  function luminance(hex: string): number {
    const full = hex.replace("#", "");
    const [r, g, b] = [0, 2, 4].map((i) => parseInt(full.slice(i, i + 2), 16) / 255).map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
  }
  function ratio(a: string, b: string): number {
    const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x);
    return (hi + 0.05) / (lo + 0.05);
  }
  const pick = (palette: ReadonlyMap<string, string>, name: string): string => {
    const value = palette.get(name);
    if (!value) throw new Error(`Thiếu biến màu --${name} trong taskpane.css`);
    return value;
  };

  const TEXT = ["text", "muted", "accent", "accent-strong", "ok", "warn", "err"];
  const SURFACES = ["bg", "surface", "surface-2", "accent-soft", "ok-bg", "warn-bg", "err-bg"];

  it("có đủ biến màu để kiểm tra", () => {
    for (const name of [...TEXT, ...SURFACES, "accent-ink", "header-bg", "header-ink"]) expect(tokens.has(name)).toBe(true);
  });

  describe.each(palettes)("giao diện %s", (_name, palette) => {
    it.each(TEXT)("chữ --%s đủ tương phản trên mọi nền của giao diện", (name) => {
      for (const surface of SURFACES) expect(ratio(pick(palette, name), pick(palette, surface))).toBeGreaterThanOrEqual(4.5);
    });

    it("chữ trên nút nền màu nhấn đủ tương phản", () => {
      expect(ratio(pick(palette, "accent-ink"), pick(palette, "accent"))).toBeGreaterThanOrEqual(4.5);
      expect(ratio(pick(palette, "accent-ink"), pick(palette, "accent-strong"))).toBeGreaterThanOrEqual(4.5);
    });
  });

  it.each(headers)("chữ trên dải đầu trang (%s) đủ tương phản", (_name, palette) => {
    expect(ratio(pick(palette, "header-ink"), pick(palette, "header-bg"))).toBeGreaterThanOrEqual(4.5);
  });

  it("CSS không có màu chữ viết cứng ngoài biến", () => {
    const hard = [...css.matchAll(/(?<![-\w])color:\s*(#[0-9a-fA-F]{3,6})/g)].map((m) => m[1]);
    expect(hard).toEqual([]);
  });

  it("có kiểu hiển thị khi điều khiển được chọn bằng bàn phím", () => {
    expect(css).toContain(":focus-visible");
  });

  it("trang khai báo ngôn ngữ tiếng Việt và có vùng cho ribbon, nút gạt chuẩn, các tab và trạng thái", () => {
    expect(html).toContain('<html lang="vi">');
    for (const id of ['id="tabs"', 'id="standard-switch"', 'id="panels"', 'id="status"']) expect(html).toContain(id);
  });
});
