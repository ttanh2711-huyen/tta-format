import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

/**
 * Quyền riêng tư (T24): mã nguồn của add-in không gọi mạng. Nội dung văn bản chỉ đi từ Word vào bộ nhớ của task pane; cấu hình
 * lưu ở localStorage của máy. Test này quét mã nguồn (trừ mã thăm dò spike chỉ chạy khi phát triển) tìm mọi cách gửi dữ liệu ra
 * ngoài; bước kiểm tra bản build (`tools/check-release.js`) kiểm tra thêm địa chỉ ngoài trong bản phát hành.
 */

const FORBIDDEN: Array<[RegExp, string]> = [
  [/\bfetch\s*\(/, "fetch"],
  [/\bXMLHttpRequest\b/, "XMLHttpRequest"],
  [/\bsendBeacon\b/, "navigator.sendBeacon"],
  [/\bWebSocket\b/, "WebSocket"],
  [/\bEventSource\b/, "EventSource"],
  [/\bimportScripts\b/, "importScripts"],
  [/\bnavigator\.serviceWorker\b/, "serviceWorker"],
  [/https?:\/\/(?!schemas\.|www\.w3\.org|purl\.org|appsforoffice\.microsoft\.com)[a-z0-9.-]+\.[a-z]{2,}/i, "địa chỉ http(s) ngoài"],
];

function sourceFiles(dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) return entry.name === "spike" ? [] : sourceFiles(path);
    return /\.(ts|html)$/.test(entry.name) && !entry.name.endsWith(".test.ts") ? [path] : [];
  });
}

describe("không gọi mạng", () => {
  const files = sourceFiles(__dirname);

  it("có mã nguồn để quét", () => {
    expect(files.length).toBeGreaterThan(50);
  });

  it.each(FORBIDDEN.map(([pattern, name]) => [name, pattern] as const))("không có %s", (_name, pattern) => {
    const hits = files.filter((file) => pattern.test(readFileSync(file, "utf8"))).map((file) => file.slice(__dirname.length + 1));
    expect(hits).toEqual([]);
  });

  it("trang chỉ nạp một script ngoài: Office.js của Microsoft", () => {
    const html = readFileSync(join(__dirname, "taskpane", "taskpane.html"), "utf8");
    const scripts = [...html.matchAll(/<script[^>]+src="([^"]+)"/g)].map((m) => m[1]);
    expect(scripts).toEqual(["https://appsforoffice.microsoft.com/lib/1/hosted/office.js"]);
  });
});
