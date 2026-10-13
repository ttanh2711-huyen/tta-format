const createConfig = require("../webpack.config.js");

/** Mẫu trang task pane phải được điền đủ chỗ trống ở mọi chế độ build (lỗi "officeScript is not defined" khi chạy dev server). */
describe("trang task pane do webpack dựng", () => {
  async function pages(env) {
    const config = await createConfig(env);
    return Object.fromEntries(
      config.plugins
        .filter((plugin) => plugin.constructor.name === "HtmlWebpackPlugin" && plugin.userOptions.templateContent)
        .map((plugin) => [plugin.userOptions.filename, plugin.userOptions.templateContent()]),
    );
  }

  it("bản phát hành: chỉ có taskpane.html, đã nhúng Office.js và kiểu dáng, không còn chỗ trống", async () => {
    const result = await pages({ release: true });
    expect(Object.keys(result)).toEqual(["taskpane.html"]);
    expect(result["taskpane.html"]).not.toContain("<%");
    expect(result["taskpane.html"]).toContain('<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>');
    expect(result["taskpane.html"]).toContain("--accent:");
  });

  it("bản phát triển: có thêm preview.html không nạp Office.js thật", async () => {
    const result = await pages({});
    expect(Object.keys(result).sort()).toEqual(["preview.html", "taskpane.html"]);
    expect(result["preview.html"]).not.toContain("<%");
    expect(result["preview.html"]).not.toContain("office.js");
    expect(result["taskpane.html"]).toContain("office.js");
  });
});
