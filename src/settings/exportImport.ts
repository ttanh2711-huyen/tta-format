import { STANDARDS } from "../standards";
import type { StandardId } from "../standards/types";
import { agencyKey, type AgencyList } from "./agencies";
import { loadDefaultStandard, saveDefaultStandard } from "./defaultStandard";
import { defaultsKey, isPreset, readUserPresets, writeUserPresets, type Preset } from "./presets";
import { allProfiles, isAgencyProfile, replaceProfiles, validateProfile, type AgencyProfile } from "./profile";
import { listCounters, recordNumber, type Counter } from "./soVanBan";
import type { KeyValueStore } from "./storage";

/**
 * Xuất và nhập cấu hình (T22): một file JSON gồm hồ sơ cơ quan, preset của người dùng, giá trị mặc định của từng chuẩn,
 * danh sách tên và chữ viết tắt cơ quan đã lưu và chuẩn mặc định. Có số phiên bản định dạng để file cũ vẫn đọc được khi
 * định dạng đổi; file có phiên bản mới hơn bản đang chạy bị từ chối thay vì đọc sai.
 */

export const CONFIG_FORMAT = "nd30-addin-config";
export const CONFIG_VERSION = 1;

export interface ConfigFile {
  format: typeof CONFIG_FORMAT;
  version: number;
  exportedAt: string;
  defaultStandard: StandardId;
  profiles: AgencyProfile[];
  presets: Preset[];
  defaults: Partial<Record<StandardId, { values: Record<string, unknown>; exceptions: string[] }>>;
  agencies: Partial<Record<StandardId, { names: string[]; abbrs: string[]; recipients: string[] }>>;
  /** Sổ số văn bản: số cuối đã dùng theo chuẩn, năm, loại (có từ T26). */
  counters?: Partial<Record<StandardId, Counter[]>>;
}

function readJson(store: KeyValueStore, key: string): unknown {
  try {
    const raw = store.get(key);
    return raw === null ? undefined : (JSON.parse(raw) as unknown);
  } catch {
    return undefined;
  }
}

const strings = (value: unknown): string[] => (Array.isArray(value) ? value.filter((item): item is string => typeof item === "string") : []);

export function exportConfig(store: KeyValueStore, now: Date = new Date()): string {
  const defaults: ConfigFile["defaults"] = {};
  const agencies: ConfigFile["agencies"] = {};
  for (const standard of STANDARDS) {
    const saved = readJson(store, defaultsKey(standard.id));
    if (typeof saved === "object" && saved !== null) {
      const { values, exceptions } = saved as { values?: unknown; exceptions?: unknown };
      defaults[standard.id] = { values: typeof values === "object" && values !== null ? (values as Record<string, unknown>) : {}, exceptions: strings(exceptions) };
    }
    const names = strings(readJson(store, agencyKey(standard.id, "names")));
    const abbrs = strings(readJson(store, agencyKey(standard.id, "abbrs")));
    const recipients = strings(readJson(store, agencyKey(standard.id, "recipients")));
    if (names.length > 0 || abbrs.length > 0 || recipients.length > 0) agencies[standard.id] = { names, abbrs, recipients };
  }
  const counters: NonNullable<ConfigFile["counters"]> = {};
  for (const standard of STANDARDS) {
    const list = listCounters(store, standard);
    if (list.length > 0) counters[standard.id] = list;
  }
  const file: ConfigFile = {
    format: CONFIG_FORMAT,
    version: CONFIG_VERSION,
    exportedAt: now.toISOString(),
    defaultStandard: loadDefaultStandard(store).id,
    profiles: allProfiles(store),
    presets: readUserPresets(store),
    defaults,
    agencies,
    counters,
  };
  return JSON.stringify(file, null, 2);
}

export interface ImportReport {
  ok: boolean;
  /** Lý do từ chối cả file (khi `ok` là false). */
  error?: string;
  profiles: number;
  presets: number;
  defaults: number;
  agencyEntries: number;
  counters: number;
  /** Mục bị bỏ qua kèm lý do. */
  skipped: string[];
}

const EMPTY: ImportReport = { ok: false, profiles: 0, presets: 0, defaults: 0, agencyEntries: 0, counters: 0, skipped: [] };

/**
 * Nhập cấu hình, gộp vào cấu hình hiện có: hồ sơ và preset trùng `id` được thay bằng bản trong file, mục mới được thêm,
 * danh sách cơ quan được hợp lại. Hồ sơ không hợp lệ với chuẩn của nó bị bỏ qua và báo lại. Không ghi gì nếu file không hợp lệ.
 */
export function importConfig(store: KeyValueStore, text: string): ImportReport {
  let data: unknown;
  try {
    data = JSON.parse(text);
  } catch {
    return { ...EMPTY, error: "File không phải JSON hợp lệ." };
  }
  if (typeof data !== "object" || data === null || (data as Record<string, unknown>).format !== CONFIG_FORMAT) {
    return { ...EMPTY, error: "Đây không phải file cấu hình của add-in này." };
  }
  const file = data as Record<string, unknown>;
  if (typeof file.version !== "number" || !Number.isInteger(file.version) || file.version < 1) return { ...EMPTY, error: "File thiếu số phiên bản định dạng." };
  if (file.version > CONFIG_VERSION) return { ...EMPTY, error: `File có định dạng phiên bản ${file.version}, mới hơn bản add-in này hiểu (${CONFIG_VERSION}); hãy cập nhật add-in.` };

  const report: ImportReport = { ...EMPTY, ok: true, skipped: [] };

  const incomingProfiles = Array.isArray(file.profiles) ? file.profiles : [];
  const profiles = new Map(allProfiles(store).map((item) => [item.id, item]));
  for (const item of incomingProfiles) {
    if (!isAgencyProfile(item)) {
      report.skipped.push("Một hồ sơ cơ quan sai cấu trúc.");
      continue;
    }
    const standard = STANDARDS.find((s) => s.id === item.standardId);
    if (!standard) {
      report.skipped.push(`Hồ sơ "${item.label}": chuẩn "${item.standardId}" không có trong add-in này.`);
      continue;
    }
    const problems = validateProfile(standard, item);
    if (problems.length > 0) {
      report.skipped.push(`Hồ sơ "${item.label}": ${problems.join(" ")}`);
      continue;
    }
    profiles.set(item.id, item);
    report.profiles += 1;
  }

  const incomingPresets = Array.isArray(file.presets) ? file.presets : [];
  const presets = new Map(readUserPresets(store).map((item) => [item.id, item]));
  for (const item of incomingPresets) {
    if (!isPreset(item)) {
      report.skipped.push("Một preset sai cấu trúc.");
      continue;
    }
    presets.set(item.id, item as Preset);
    report.presets += 1;
  }

  const defaultsIn = typeof file.defaults === "object" && file.defaults !== null ? (file.defaults as Record<string, unknown>) : {};
  const defaultsOut: Array<[StandardId, unknown]> = [];
  for (const standard of STANDARDS) {
    const value = defaultsIn[standard.id];
    if (typeof value === "object" && value !== null) {
      defaultsOut.push([standard.id, value]);
      report.defaults += 1;
    }
  }

  const agenciesIn = typeof file.agencies === "object" && file.agencies !== null ? (file.agencies as Record<string, unknown>) : {};
  const agencyOut: Array<[StandardId, AgencyList, string[]]> = [];
  for (const standard of STANDARDS) {
    const entry = agenciesIn[standard.id];
    if (typeof entry !== "object" || entry === null) continue;
    for (const list of ["names", "abbrs", "recipients"] as const) {
      const incoming = strings((entry as Record<string, unknown>)[list]);
      if (incoming.length === 0) continue;
      const current = strings(readJson(store, agencyKey(standard.id, list)));
      const merged = [...current, ...incoming.filter((name) => !current.some((c) => c.toLowerCase() === name.toLowerCase()))];
      agencyOut.push([standard.id, list, merged]);
      report.agencyEntries += merged.length - current.length;
    }
  }

  const countersIn = typeof file.counters === "object" && file.counters !== null ? (file.counters as Record<string, unknown>) : {};
  const counterOut: Array<[(typeof STANDARDS)[number], Counter]> = [];
  for (const standard of STANDARDS) {
    const list = countersIn[standard.id];
    if (!Array.isArray(list)) continue;
    for (const item of list) {
      const c = item as Partial<Counter> | null;
      if (c && typeof c.year === "number" && typeof c.typeKey === "string" && typeof c.last === "number") {
        counterOut.push([standard, { year: c.year, typeKey: c.typeKey, last: c.last }]);
        report.counters += 1;
      } else {
        report.skipped.push("Một mục sổ số sai cấu trúc.");
      }
    }
  }

  replaceProfiles(store, [...profiles.values()]);
  writeUserPresets(store, [...presets.values()]);
  for (const [id, value] of defaultsOut) store.set(defaultsKey(id), JSON.stringify(value));
  for (const [id, list, merged] of agencyOut) store.set(agencyKey(id, list), JSON.stringify(merged));
  for (const [standard, c] of counterOut) recordNumber(store, standard, c.year, c.typeKey === "-" ? null : c.typeKey, c.last);
  if (STANDARDS.some((s) => s.id === file.defaultStandard)) saveDefaultStandard(store, file.defaultStandard as StandardId);
  return report;
}
