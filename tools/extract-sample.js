/**
 * Trích file mẫu .docx thành fixture ẩn danh (T2A): chỉ giữ định dạng, chữ thể thức công khai và nhãn đầu dòng.
 * Dùng: node tools/extract-sample.js "<đường dẫn .docx>" [fixtures/congvan-mau.json]
 * File .docx gốc nằm ngoài repo và không bao giờ được sao chép vào repo.
 */
const fs = require("node:fs");
const path = require("node:path");
const AdmZip = require("adm-zip");
const ts = require("typescript");

// Nạp trực tiếp mã TypeScript của src/rules mà không cần bước build riêng.
require.extensions[".ts"] = (module, filename) => {
  const { outputText } = ts.transpileModule(fs.readFileSync(filename, "utf8"), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 } });
  module._compile(outputText, filename);
};
const { anonymizeModel, extractDocModel } = require("../src/rules/docModel.ts");

function buildFixture(docxPath) {
  const zip = new AdmZip(docxPath);
  const text = (name) => zip.readAsText(name, "utf8");
  const headerXmls = zip
    .getEntries()
    .map((entry) => entry.entryName)
    .filter((name) => /^word\/header\d+\.xml$/.test(name))
    .sort()
    .map(text);
  return anonymizeModel(extractDocModel({ documentXml: text("word/document.xml"), headerXmls }));
}

module.exports = { buildFixture };

if (require.main === module) {
  const [source, target = path.join("fixtures", "congvan-mau.json")] = process.argv.slice(2);
  if (!source) {
    process.stderr.write('Dùng: node tools/extract-sample.js "<file .docx>" [fixtures/congvan-mau.json]\n');
    process.exit(1);
  }
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, `${JSON.stringify(buildFixture(source), null, 2)}\n`, "utf8");
  process.stdout.write(`Đã ghi ${target}\n`);
}
