const { scanRelease } = require("./check-release");

const manifest = { path: "manifest.xml", content: '<SourceLocation DefaultValue="https://tenban.github.io/nd30-word-addin/taskpane.html"/>' };
const good = [manifest, { path: "taskpane.html", content: '<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>' }, { path: "taskpane.js", content: 'const ns="http://schemas.openxmlformats.org/wordprocessingml/2006/main";' }, { path: "assets/icon-16.png", content: null }];

describe("kiểm tra bản phát hành", () => {
  it("bản đúng thì không lỗi", () => {
    expect(scanRelease(good)).toEqual([]);
  });

  it("chấp nhận tên JS có mã băm và địa chỉ trang kèm phiên bản", () => {
    const versioned = { path: "manifest.xml", content: '<SourceLocation DefaultValue="https://tenban.github.io/nd30-word-addin/taskpane.html?v=1.0.7"/>' };
    expect(scanRelease([versioned, ...good.slice(1), { path: "taskpane.1a2b3c4d.js", content: "" }])).toEqual([]);
    expect(scanRelease([...good, { path: "taskpane.evil.js", content: "" }]).join(" ")).toContain("không mong đợi");
  });

  it("thiếu manifest", () => {
    expect(scanRelease(good.filter((f) => f.path !== "manifest.xml")).join(" ")).toContain("Thiếu manifest.xml");
  });

  it("manifest còn localhost hoặc không trỏ về https", () => {
    expect(scanRelease([{ path: "manifest.xml", content: '<SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>' }]).join(" ")).toContain("localhost");
    expect(scanRelease([{ path: "manifest.xml", content: "<OfficeApp/>" }]).join(" ")).toContain("thiếu địa chỉ taskpane.html");
  });

  it("file không mong đợi (trang thăm dò, tài liệu, mã nguồn) bị báo", () => {
    for (const path of ["spike.html", "docs/test-matrix.md", "src/a.ts", "fixtures/congvan-mau.json"]) {
      expect(scanRelease([...good, { path, content: "" }]).join(" ")).toContain(path);
    }
  });

  it("địa chỉ ngoài không được phép bị báo", () => {
    const problems = scanRelease([...good, { path: "taskpane.js", content: 'fetch("https://analytics.example.com/x")' }]);
    expect(problems.join(" ")).toContain("analytics.example.com");
  });

  it("dấu vết dữ liệu cá nhân bị báo", () => {
    for (const content of ["C:\\Users\\nguoi\\dev", "/Users/nguoi/dev/x", "lien-he nguoi@gmail.com"]) {
      expect(scanRelease([...good, { path: "taskpane.js", content }]).join(" ")).toContain("dữ liệu cá nhân");
    }
  });
});
