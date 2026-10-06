export type ProbeStatus = "pass" | "fail" | "unsupported" | "manual" | "skipped";
export type ProbeGroup = "chung" | "nd30" | "hd05";

export interface ProbeResult {
  id: string;
  title: string;
  group: ProbeGroup;
  /** Mục mà thiếu nó thì không thể làm tiếp theo hướng Office.js. */
  critical: boolean;
  status: ProbeStatus;
  detail: string;
  durationMs: number;
  data?: unknown;
}

export interface EnvironmentInfo {
  host: string;
  platform: string;
  version: string;
  /** Ví dụ "WordApiDesktop 1.3": true. */
  sets: Record<string, boolean>;
}

export interface SpikeReport {
  generatedAt: string;
  environment: EnvironmentInfo;
  results: ProbeResult[];
}

export function summarize(results: readonly ProbeResult[]): Record<ProbeStatus, number> {
  const counts: Record<ProbeStatus, number> = { pass: 0, fail: 0, unsupported: 0, manual: 0, skipped: 0 };
  for (const result of results) {
    counts[result.status] += 1;
  }
  return counts;
}

export type Verdict = "go" | "go-with-fallback" | "no-go";

const NOT_OK: readonly ProbeStatus[] = ["fail", "unsupported"];

/** Mục bắt buộc không đạt thì no-go; chỉ mục khác không đạt thì go kèm phương án thay thế. */
export function goNoGo(results: readonly ProbeResult[]): { verdict: Verdict; reasons: string[] } {
  const label = (r: ProbeResult) => `${r.title}: ${r.status}`;
  const criticalBad = results.filter((r) => r.critical && NOT_OK.includes(r.status));
  if (criticalBad.length > 0) {
    return { verdict: "no-go", reasons: criticalBad.map(label) };
  }
  const otherBad = results.filter((r) => NOT_OK.includes(r.status));
  if (otherBad.length > 0) {
    return { verdict: "go-with-fallback", reasons: otherBad.map(label) };
  }
  return { verdict: "go", reasons: [] };
}

function cellText(text: string): string {
  return text.replace(/\r?\n/g, " ").replace(/\|/g, "\\|");
}

export function toMarkdown(report: SpikeReport): string {
  const { environment, results } = report;
  const counts = summarize(results);
  const { verdict, reasons } = goNoGo(results);
  const sets = Object.entries(environment.sets).map(([name, ok]) => `- ${name}: ${ok ? "có" : "không"}`);
  const rows = results.map(
    (r) => `| ${r.id} | ${cellText(r.title)} | ${r.group} | ${r.status} | ${cellText(r.detail)} |`,
  );
  return [
    "# Kết quả spike Office.js",
    "",
    `Chạy lúc: ${report.generatedAt}`,
    "",
    "## Môi trường",
    "",
    `- Ứng dụng: ${environment.host}`,
    `- Nền tảng: ${environment.platform}`,
    `- Phiên bản: ${environment.version}`,
    ...sets,
    "",
    "## Tổng hợp",
    "",
    `- pass: ${counts.pass}, fail: ${counts.fail}, unsupported: ${counts.unsupported}, manual: ${counts.manual}, skipped: ${counts.skipped}`,
    `- **Kết luận: ${verdict}**${reasons.length > 0 ? ` (${reasons.join("; ")})` : ""}`,
    "",
    "## Chi tiết",
    "",
    "| Mã | Nội dung kiểm tra | Nhóm | Kết quả | Chi tiết |",
    "| --- | --- | --- | --- | --- |",
    ...rows,
    "",
  ].join("\n");
}
