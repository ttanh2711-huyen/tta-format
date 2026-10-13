import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describeDegraded, describeMissing, evaluateFeatures, FEATURES, type IsSetSupported } from "./capabilities";

const supports = (...sets: string[]): IsSetSupported => (set, version) => sets.includes(`${set} ${version}`);

describe("đánh giá tính năng theo bộ API", () => {
  it("Word đủ bộ API: mọi tính năng chạy, không cảnh báo", () => {
    const all = supports("WordApi 1.3", "WordApi 1.6", "WordApiDesktop 1.1", "WordApiDesktop 1.3", "WordApiDesktop 1.4");
    for (const status of evaluateFeatures(all)) {
      expect(status.supported).toBe(true);
      expect(describeMissing(status)).toBe("");
      expect(describeDegraded(status)).toBe("");
    }
  });

  it("Word cũ (chỉ WordApi 1.3): tính năng cần bộ mới bị tắt kèm lý do tiếng Việt nêu bản Word", () => {
    const statuses = evaluateFeatures(supports("WordApi 1.3"));
    const byId = Object.fromEntries(statuses.map((s) => [s.feature.id, s]));
    expect(byId.insertBlocks.supported).toBe(true);
    expect(byId.documentStandard.supported).toBe(true);
    for (const id of ["pageSetup", "styles", "pageNumber", "check", "createDocument", "formatBody"]) {
      expect(byId[id].supported).toBe(false);
      const text = describeMissing(byId[id]);
      expect(text).toContain("Microsoft 365 bản");
      expect(text).toContain("bị tắt");
    }
    expect(describeMissing(byId.pageSetup)).toContain("WordApiDesktop 1.3 (Microsoft 365 bản 2507 trở lên)");
  });

  it("thiếu bộ API tuỳ chọn: tính năng vẫn chạy nhưng có cảnh báo dùng cách thay thế", () => {
    const status = evaluateFeatures(supports("WordApi 1.3", "WordApi 1.6", "WordApiDesktop 1.3")).find((s) => s.feature.id === "formatBody")!;
    expect(status.supported).toBe(true);
    expect(status.missingOptional.map((r) => `${r.set} ${r.version}`)).toEqual(["WordApiDesktop 1.4"]);
    expect(describeDegraded(status)).toContain("WordApiDesktop 1.4");
  });

  it("tính năng không chạy được không báo 'chạy một phần'", () => {
    const status = evaluateFeatures(supports())[0];
    expect(describeDegraded(status)).toBe("");
  });

  it("mỗi tính năng có mã duy nhất và nhãn", () => {
    expect(new Set(FEATURES.map((f) => f.id)).size).toBe(FEATURES.length);
    expect(FEATURES.every((f) => f.label !== "" && f.requires.length > 0)).toBe(true);
  });
});

describe("bảng tính năng khớp với mã nguồn", () => {
  const known = new Set(FEATURES.flatMap((f) => [...f.requires, ...(f.optional ?? [])].map((r) => `${r.set} ${r.version}`)));

  function sourceFiles(dir: string): string[] {
    return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
      const path = join(dir, entry.name);
      if (entry.isDirectory()) return entry.name === "selftest" ? [] : sourceFiles(path);
      return entry.name.endsWith(".ts") && !entry.name.endsWith(".test.ts") && entry.name !== "capabilities.ts" ? [path] : [];
    });
  }

  it("mọi yêu cầu isSetSupported trong mã đều có trong bảng", () => {
    const used = new Set<string>();
    for (const file of sourceFiles(join(__dirname, ".."))) {
      for (const match of readFileSync(file, "utf8").matchAll(/isSetSupported\("(WordApi(?:Desktop)?)",\s*"([\d.]+)"\)/g)) used.add(`${match[1]} ${match[2]}`);
    }
    expect(used.size).toBeGreaterThan(0);
    const missing = [...used].filter((item) => !known.has(item));
    expect(missing).toEqual([]);
  });
});
