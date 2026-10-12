const { readFileSync } = require("node:fs");
const { join } = require("node:path");
const { buildManifest, normalizeBase, officeVersion } = require("./make-manifest");

const source = readFileSync(join(__dirname, "..", "manifest.xml"), "utf8");
const BASE = "https://tenban.github.io/nd30-word-addin";

describe("manifest phát hành", () => {
  it("không còn localhost; mọi địa chỉ trỏ về trang phát hành", () => {
    const xml = buildManifest(source, BASE, "1.0.0");
    expect(xml).not.toContain("localhost");
    expect(xml).toContain(`<SourceLocation DefaultValue="${BASE}/taskpane.html"/>`);
    expect(xml).toContain(`DefaultValue="${BASE}/assets/icon-32.png"`);
    expect(xml).toContain("<AppDomain>https://tenban.github.io</AppDomain>");
  });

  it("số phiên bản lấy từ package.json, dạng bốn số", () => {
    expect(buildManifest(source, BASE, "1.2.3")).toContain("<Version>1.2.3.0</Version>");
    expect(officeVersion("1.0.0")).toBe("1.0.0.0");
    expect(officeVersion("2.0.1-rc.1")).toBe("2.0.1.0");
  });

  it("giữ nguyên Id và các phần còn lại của manifest", () => {
    const xml = buildManifest(source, BASE, "1.0.0");
    expect(xml).toContain("<Id>f47a9dbd-3e89-4ae6-a70d-f52558640c57</Id>");
    expect(xml).toContain("<Permissions>ReadWriteDocument</Permissions>");
    expect(xml.match(/<bt:Image /g)?.length).toBe(source.match(/<bt:Image /g)?.length);
  });

  it("chấp nhận địa chỉ có dấu gạch chéo cuối", () => {
    expect(normalizeBase(`${BASE}/`)).toBe(BASE);
    expect(buildManifest(source, `${BASE}/`, "1.0.0")).toContain(`${BASE}/taskpane.html`);
  });

  it.each([["http://tenban.github.io/x"], ["https://localhost:4000"], ["tenban.github.io"], ["https://có khoảng trắng.vn"], [""]])("từ chối địa chỉ sai %j", (bad) => {
    expect(() => buildManifest(source, bad, "1.0.0")).toThrow();
  });

  it("từ chối phiên bản dưới 1.0.0 vì Office không chấp nhận", () => {
    expect(() => buildManifest(source, BASE, "0.9.0")).toThrow("từ 1.0.0");
  });

  it("từ chối số phiên bản sai và manifest không có chỗ để thay", () => {
    expect(() => buildManifest(source, BASE, "abc")).toThrow();
    expect(() => buildManifest("<OfficeApp/>", BASE, "1.0.0")).toThrow();
  });

  it("tên add-in là TTA-Format ở tên hiển thị, thẻ ribbon và nút", () => {
    expect(source).toContain('<DisplayName DefaultValue="TTA-Format"/>');
    expect(source).toContain('<bt:String id="Tab.Label" DefaultValue="TTA-Format"/>');
    expect(source).toContain('<bt:String id="Button.Label" DefaultValue="Mở TTA-Format"/>');
    expect(source).not.toMatch(/NĐ 30 -|Mở NĐ 30|DefaultValue="NĐ 30"/);
  });
});
