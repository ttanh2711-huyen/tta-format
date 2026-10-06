import { PROBES, type Probe } from "./probes";
import type { EnvironmentInfo, ProbeResult } from "./report";
import { probesForRunAll } from "./selection";

const WORD_API_VERSIONS = ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9"];
const WORD_API_DESKTOP_VERSIONS = ["1.1", "1.2", "1.3", "1.4", "1.5"];

export function collectEnvironment(): EnvironmentInfo {
  const diagnostics = Office.context.diagnostics;
  const sets: Record<string, boolean> = {};
  for (const version of WORD_API_VERSIONS) {
    sets[`WordApi ${version}`] = Office.context.requirements.isSetSupported("WordApi", version);
  }
  for (const version of WORD_API_DESKTOP_VERSIONS) {
    sets[`WordApiDesktop ${version}`] = Office.context.requirements.isSetSupported("WordApiDesktop", version);
  }
  return {
    host: String(diagnostics?.host ?? "không rõ"),
    platform: String(diagnostics?.platform ?? "không rõ"),
    version: String(diagnostics?.version ?? "không rõ"),
    sets,
  };
}

function describeError(error: unknown): string {
  if (error instanceof Error) {
    const detailed = error as Error & { code?: string; debugInfo?: unknown };
    const parts = [detailed.code ? `${detailed.code}: ${error.message}` : error.message];
    if (detailed.debugInfo) {
      parts.push(`debugInfo=${JSON.stringify(detailed.debugInfo)}`);
    }
    return parts.join(" ");
  }
  return String(error);
}

function missingRequirements(probe: Probe): string[] {
  return (probe.requires ?? [])
    .filter((r) => !Office.context.requirements.isSetSupported(r.set, r.version))
    .map((r) => `${r.set} ${r.version}`);
}

export async function runProbe(probe: Probe): Promise<ProbeResult> {
  const base = { id: probe.id, title: probe.title, group: probe.group, critical: probe.critical };
  const missing = missingRequirements(probe);
  if (missing.length > 0) {
    return { ...base, status: "unsupported", detail: `Thiếu requirement set: ${missing.join(", ")}.`, durationMs: 0 };
  }
  const started = performance.now();
  try {
    const outcome = await probe.run();
    return {
      ...base,
      status: outcome.status,
      detail: outcome.detail,
      data: outcome.data,
      durationMs: Math.round(performance.now() - started),
    };
  } catch (error) {
    return {
      ...base,
      status: "fail",
      detail: `Lỗi khi chạy: ${describeError(error)}`,
      durationMs: Math.round(performance.now() - started),
    };
  }
}

/**
 * Chạy các phép thử không rủi ro lần lượt. `onBefore` được gọi trước mỗi phép thử để lưu mã đang
 * chạy; nếu Word đóng giữa chừng thì báo cáo đã lưu cho biết phép thử nào là thủ phạm.
 */
export async function runAll(
  onBefore: (probe: Probe) => Promise<void>,
  onResult: (result: ProbeResult) => void,
): Promise<ProbeResult[]> {
  const results: ProbeResult[] = [];
  for (const probe of probesForRunAll(PROBES)) {
    await onBefore(probe);
    const result = await runProbe(probe);
    results.push(result);
    onResult(result);
  }
  return results;
}
