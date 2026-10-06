import { goNoGo, summarize, toMarkdown, type ProbeResult, type SpikeReport } from "./report";

function result(partial: Partial<ProbeResult> & Pick<ProbeResult, "id" | "status">): ProbeResult {
  return {
    title: `Thử ${partial.id}`,
    group: "chung",
    critical: false,
    detail: "",
    durationMs: 5,
    ...partial,
  };
}

describe("summarize", () => {
  it("đếm số kết quả theo từng trạng thái", () => {
    const counts = summarize([
      result({ id: "a", status: "pass" }),
      result({ id: "b", status: "pass" }),
      result({ id: "c", status: "fail" }),
      result({ id: "d", status: "unsupported" }),
      result({ id: "e", status: "manual" }),
    ]);
    expect(counts).toEqual({ pass: 2, fail: 1, unsupported: 1, manual: 1, skipped: 0 });
  });
});

describe("goNoGo", () => {
  it("go khi mọi mục bắt buộc đều đạt và mục không bắt buộc đạt", () => {
    const verdict = goNoGo([
      result({ id: "a", status: "pass", critical: true }),
      result({ id: "b", status: "pass" }),
    ]);
    expect(verdict.verdict).toBe("go");
    expect(verdict.reasons).toEqual([]);
  });

  it("go-with-fallback khi chỉ mục không bắt buộc không đạt", () => {
    const verdict = goNoGo([
      result({ id: "a", status: "pass", critical: true }),
      result({ id: "b", status: "fail", title: "Thử b" }),
    ]);
    expect(verdict.verdict).toBe("go-with-fallback");
    expect(verdict.reasons).toEqual(["Thử b: fail"]);
  });

  it("no-go khi một mục bắt buộc không đạt hoặc không được hỗ trợ", () => {
    const failed = goNoGo([result({ id: "a", status: "fail", critical: true, title: "Thử a" })]);
    expect(failed.verdict).toBe("no-go");
    expect(failed.reasons).toEqual(["Thử a: fail"]);

    const unsupported = goNoGo([result({ id: "a", status: "unsupported", critical: true, title: "Thử a" })]);
    expect(unsupported.verdict).toBe("no-go");
  });

  it("mục thủ công và mục bỏ qua không làm đổi kết luận", () => {
    const verdict = goNoGo([
      result({ id: "a", status: "pass", critical: true }),
      result({ id: "b", status: "manual" }),
      result({ id: "c", status: "skipped" }),
    ]);
    expect(verdict.verdict).toBe("go");
  });
});

describe("toMarkdown", () => {
  const report: SpikeReport = {
    generatedAt: "2026-10-06T10:00:00.000Z",
    environment: { host: "Word", platform: "PC", version: "16.0.19029.20208", sets: { "WordApiDesktop 1.3": true, "WordApiDesktop 1.4": false } },
    results: [
      result({ id: "p02", title: "Lề trang", group: "nd30", status: "pass", detail: "Lề đọc lại đúng", critical: true }),
      result({ id: "p10", title: "Style", group: "chung", status: "fail", detail: "dòng một\ndòng hai | có gạch đứng" }),
    ],
  };

  it("có bảng kết quả với tiêu đề tiếng Việt, mỗi probe một dòng", () => {
    const md = toMarkdown(report);
    expect(md).toContain("| Mã | Nội dung kiểm tra | Nhóm | Kết quả | Chi tiết |");
    expect(md).toContain("| p02 | Lề trang | nd30 | pass | Lề đọc lại đúng |");
  });

  it("thoát ký tự gạch đứng và xuống dòng trong chi tiết để không vỡ bảng", () => {
    const md = toMarkdown(report);
    expect(md).toContain("dòng một dòng hai \\| có gạch đứng");
    expect(md.split("\n").filter((line) => line.startsWith("| p10"))).toHaveLength(1);
  });

  it("ghi môi trường và các requirement set được hỗ trợ", () => {
    const md = toMarkdown(report);
    expect(md).toContain("16.0.19029.20208");
    expect(md).toContain("WordApiDesktop 1.3: có");
    expect(md).toContain("WordApiDesktop 1.4: không");
  });

  it("ghi kết luận go/no-go", () => {
    const md = toMarkdown(report);
    expect(md).toMatch(/Kết luận: go-with-fallback/);
  });
});

describe("toMarkdown khi chạy bị dừng đột ngột", () => {
  const base: SpikeReport = {
    generatedAt: "2026-10-06T10:00:00.000Z",
    environment: { host: "Word", platform: "PC", version: "16.0", sets: {} },
    results: [result({ id: "p02", status: "pass", critical: true })],
  };

  it("ghi rõ mục đang chạy lúc dừng", () => {
    const md = toMarkdown({ ...base, running: "p12" });
    expect(md).toContain("Dừng đột ngột khi đang chạy: p12");
  });

  it("không có dòng đó khi chạy xong bình thường", () => {
    expect(toMarkdown(base)).not.toContain("Dừng đột ngột");
  });
});
