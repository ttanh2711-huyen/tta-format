/**
 * Kiểm tra bản phát hành (T24) trước khi đưa lên trang công khai: thư mục dist chỉ gồm các file mong đợi, không chứa dấu
 * vết dữ liệu cá nhân (đường dẫn máy, địa chỉ thư), không có địa chỉ ngoài trừ thư viện Office.js của Microsoft, không còn
 * localhost trong manifest, và manifest có địa chỉ phát hành hợp lệ.
 *
 * Dùng: node tools/check-release.js [thư mục, mặc định dist]
 */
const { readdirSync, readFileSync, statSync } = require("node:fs");
const { join, relative, resolve } = require("node:path");

/** Tên file hoặc thư mục được phép ở gốc bản phát hành. */
const ALLOWED_ROOT = new Set(["taskpane.html", "taskpane.js", "manifest.xml", "assets"]);
/** Địa chỉ ngoài được phép xuất hiện trong bản build: thư viện Office.js và các địa chỉ định danh XML/tài liệu chuẩn. */
const ALLOWED_HOSTS = new Set(["appsforoffice.microsoft.com", "schemas.microsoft.com", "schemas.openxmlformats.org", "www.w3.org", "purl.org", "schemas.xmlsoap.org"]);
/** Dấu vết dữ liệu cá nhân không được có trong bản build. */
const PERSONAL = [/[A-Za-z]:\\Users\\/, /\/Users\/[A-Za-z0-9._-]+\//, /\b[\w.+-]+@(?:gmail|yahoo|outlook|hotmail)\.com\b/i];

/** files: [{ path: đường dẫn tương đối dùng dấu /, content: chuỗi hoặc null nếu là file nhị phân }]. Trả về danh sách lỗi. */
function scanRelease(files) {
  const problems = [];
  const manifest = files.find((file) => file.path === "manifest.xml");
  if (!manifest) problems.push("Thiếu manifest.xml (chạy npm run manifest:release trước).");
  for (const file of files) {
    const top = file.path.split("/")[0];
    if (!ALLOWED_ROOT.has(top)) problems.push(`File không mong đợi trong bản phát hành: ${file.path}`);
    if (file.content === null) continue;
    for (const pattern of PERSONAL) {
      if (pattern.test(file.content)) problems.push(`${file.path}: có dấu vết dữ liệu cá nhân (${pattern}).`);
    }
    for (const match of file.content.matchAll(/https?:\/\/([a-z0-9.-]+)/gi)) {
      const host = match[1].toLowerCase();
      if (file.path === "manifest.xml") continue;
      if (!ALLOWED_HOSTS.has(host)) problems.push(`${file.path}: địa chỉ ngoài không được phép: ${match[0]}`);
    }
  }
  if (manifest && manifest.content) {
    if (/localhost/.test(manifest.content)) problems.push("manifest.xml còn localhost.");
    if (!/<SourceLocation DefaultValue="https:\/\/[^"]+\/taskpane\.html"\/>/.test(manifest.content)) problems.push("manifest.xml thiếu địa chỉ taskpane.html dạng https.");
  }
  return problems;
}

function collect(dir, root = dir) {
  return readdirSync(dir).flatMap((name) => {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) return collect(path, root);
    const rel = relative(root, path).split("\\").join("/");
    return [{ path: rel, content: /\.(html|js|xml|json|css|txt|svg)$/i.test(name) ? readFileSync(path, "utf8") : null }];
  });
}

if (require.main === module) {
  const dir = resolve(process.argv[2] ?? "dist");
  const problems = scanRelease(collect(dir));
  if (problems.length > 0) {
    console.error(problems.join("\n"));
    process.exit(1);
  }
  console.log(`Bản phát hành ${dir} đạt kiểm tra.`);
}

module.exports = { scanRelease, ALLOWED_HOSTS };
