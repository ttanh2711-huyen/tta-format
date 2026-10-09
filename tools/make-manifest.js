/**
 * Tạo manifest phát hành (T24) từ manifest.xml dùng khi phát triển: thay địa chỉ https://localhost:3000 bằng địa chỉ
 * trang phát hành, đặt số phiên bản theo package.json (dạng bốn số của Office), ghi ra thư mục dist.
 *
 * Dùng: node tools/make-manifest.js <địa chỉ gốc, ví dụ https://tenban.github.io/nd30-word-addin> [thư mục ra, mặc định dist]
 */
const { mkdirSync, readFileSync, writeFileSync } = require("node:fs");
const { join, resolve } = require("node:path");

const DEV_ORIGIN = "https://localhost:3000";

/** "0.9.0" thành "0.9.0.0"; Office yêu cầu đúng bốn số nguyên không âm. */
function officeVersion(packageVersion) {
  const parts = String(packageVersion).split(/[.+-]/).slice(0, 3);
  if (parts.length < 3 || parts.some((part) => !/^\d+$/.test(part))) {
    throw new Error(`Số phiên bản "${packageVersion}" phải có dạng x.y.z`);
  }
  return `${parts.join(".")}.0`;
}

function normalizeBase(baseUrl) {
  if (!/^https:\/\/[^\s/]+(\/[^\s]*)?$/.test(baseUrl)) {
    throw new Error(`Địa chỉ phát hành phải bắt đầu bằng https:// và không có khoảng trắng: "${baseUrl}"`);
  }
  if (baseUrl.includes("localhost")) {
    throw new Error("Địa chỉ phát hành không được là localhost.");
  }
  return baseUrl.replace(/\/+$/, "");
}

/** Trả về nội dung manifest phát hành. Ném lỗi nếu địa chỉ sai hoặc manifest gốc không có chỗ để thay. */
function buildManifest(sourceXml, baseUrl, packageVersion) {
  const base = normalizeBase(baseUrl);
  const version = officeVersion(packageVersion);
  if (!sourceXml.includes(DEV_ORIGIN)) {
    throw new Error(`manifest.xml không chứa ${DEV_ORIGIN} để thay.`);
  }
  const appDomain = new URL(base).origin;
  let xml = sourceXml.replace(/<AppDomain>[^<]*<\/AppDomain>/, `<AppDomain>${appDomain}</AppDomain>`);
  xml = xml.split(DEV_ORIGIN).join(base);
  xml = xml.replace(/<Version>[^<]*<\/Version>/, `<Version>${version}</Version>`);
  if (xml.includes("localhost")) {
    throw new Error("Manifest phát hành vẫn còn localhost.");
  }
  return xml;
}

function main(argv) {
  const [baseUrl, outDir = "dist"] = argv;
  if (!baseUrl) {
    throw new Error("Thiếu địa chỉ phát hành. Dùng: node tools/make-manifest.js https://tenban.github.io/nd30-word-addin");
  }
  const root = resolve(__dirname, "..");
  const pkg = JSON.parse(readFileSync(join(root, "package.json"), "utf8"));
  const xml = buildManifest(readFileSync(join(root, "manifest.xml"), "utf8"), baseUrl, pkg.version);
  const out = resolve(root, outDir);
  mkdirSync(out, { recursive: true });
  writeFileSync(join(out, "manifest.xml"), xml, "utf8");
  console.log(`Đã ghi ${join(out, "manifest.xml")} (phiên bản ${officeVersion(pkg.version)}, địa chỉ ${normalizeBase(baseUrl)}).`);
}

if (require.main === module) {
  try {
    main(process.argv.slice(2));
  } catch (error) {
    console.error(error.message);
    process.exit(1);
  }
}

module.exports = { buildManifest, officeVersion, normalizeBase };
