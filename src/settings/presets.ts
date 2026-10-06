import { STANDARDS } from "../standards";
import type { StandardId, StandardProfile } from "../standards/types";
import { buildRegistry, initialState, reconcile, type FormatRegistry, type FormatState, type OptionValue } from "./formatOptions";
import type { KeyValueStore } from "./storage";

/** Bộ giá trị tuỳ chọn có tên, thuộc một chuẩn. */
export interface Preset {
  id: string;
  name: string;
  standardId: StandardId;
  values: Record<string, OptionValue>;
  exceptions: string[];
  /** Preset có sẵn của chuẩn: không đổi tên, không xoá. */
  builtIn?: true;
}

export type PresetResult = { ok: true; preset: Preset } | { ok: false; reason: string };
export type SimpleResult = { ok: true } | { ok: false; reason: string };

const PRESETS_KEY = "nd30.presets";
const defaultsKey = (standardId: StandardId): string => `nd30.defaults.${standardId}`;
const BUILT_IN_REASON = "Preset có sẵn của chuẩn, không thể đổi tên hoặc xoá. Hãy nhân bản để chỉnh sửa.";

function readJson(store: KeyValueStore, key: string): unknown {
  const raw = store.get(key);
  if (raw === null) {
    return undefined;
  }
  try {
    return JSON.parse(raw);
  } catch {
    return undefined;
  }
}

function isPreset(value: unknown): value is Preset {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const p = value as Record<string, unknown>;
  return (
    typeof p.id === "string" &&
    typeof p.name === "string" &&
    STANDARDS.some((standard) => standard.id === p.standardId) &&
    typeof p.values === "object" &&
    p.values !== null &&
    Array.isArray(p.exceptions)
  );
}

function readUserPresets(store: KeyValueStore): Preset[] {
  const parsed = readJson(store, PRESETS_KEY);
  return Array.isArray(parsed) ? parsed.filter(isPreset) : [];
}

function writeUserPresets(store: KeyValueStore, presets: Preset[]): void {
  store.set(PRESETS_KEY, JSON.stringify(presets));
}

function builtInPreset(profile: StandardProfile): Preset {
  const state = initialState(buildRegistry(profile));
  return { id: `builtin:${profile.id}`, name: profile.defaultPresetName, standardId: profile.id, values: state.values, exceptions: [], builtIn: true };
}

/** Preset của một chuẩn: preset có sẵn trước, rồi preset của người dùng theo thứ tự lưu. */
export function listPresets(store: KeyValueStore, profile: StandardProfile): Preset[] {
  return [builtInPreset(profile), ...readUserPresets(store).filter((p) => p.standardId === profile.id)];
}

function nextId(all: readonly Preset[]): string {
  const used = all.map((p) => Number(/^u(\d+)$/.exec(p.id)?.[1] ?? 0));
  return `u${Math.max(0, ...used) + 1}`;
}

function nameProblem(name: string, taken: readonly string[]): string | null {
  const trimmed = name.trim();
  if (trimmed === "") {
    return "Hãy nhập tên cho preset.";
  }
  if (taken.some((t) => t.trim().toLowerCase() === trimmed.toLowerCase())) {
    return `Đã có preset tên "${trimmed}" trong chuẩn này, tên không được trùng.`;
  }
  return null;
}

export function savePreset(store: KeyValueStore, profile: StandardProfile, name: string, state: FormatState): PresetResult {
  const mine = listPresets(store, profile);
  const problem = nameProblem(name, mine.map((p) => p.name));
  if (problem) {
    return { ok: false, reason: problem };
  }
  const all = readUserPresets(store);
  const preset: Preset = { id: nextId(all), name: name.trim(), standardId: profile.id, values: { ...state.values }, exceptions: [...state.exceptions] };
  writeUserPresets(store, [...all, preset]);
  return { ok: true, preset };
}

export function renamePreset(store: KeyValueStore, id: string, name: string): SimpleResult {
  if (id.startsWith("builtin:")) {
    return { ok: false, reason: BUILT_IN_REASON };
  }
  const all = readUserPresets(store);
  const target = all.find((p) => p.id === id);
  if (!target) {
    return { ok: false, reason: "Không tìm thấy preset." };
  }
  const siblings = all.filter((p) => p.standardId === target.standardId && p.id !== id).map((p) => p.name);
  const problem = nameProblem(name, siblings);
  if (problem) {
    return { ok: false, reason: problem };
  }
  writeUserPresets(store, all.map((p) => (p.id === id ? { ...p, name: name.trim() } : p)));
  return { ok: true };
}

export function deletePreset(store: KeyValueStore, id: string): SimpleResult {
  if (id.startsWith("builtin:")) {
    return { ok: false, reason: BUILT_IN_REASON };
  }
  const all = readUserPresets(store);
  if (!all.some((p) => p.id === id)) {
    return { ok: false, reason: "Không tìm thấy preset." };
  }
  writeUserPresets(store, all.filter((p) => p.id !== id));
  return { ok: true };
}

export function duplicatePreset(store: KeyValueStore, profile: StandardProfile, id: string): PresetResult {
  const mine = listPresets(store, profile);
  const source = mine.find((p) => p.id === id);
  if (!source) {
    return { ok: false, reason: "Không tìm thấy preset." };
  }
  const taken = mine.map((p) => p.name.toLowerCase());
  let name = `${source.name} (bản sao)`;
  for (let n = 2; taken.includes(name.toLowerCase()); n += 1) {
    name = `${source.name} (bản sao ${n})`;
  }
  return savePreset(store, profile, name, { values: source.values, exceptions: source.exceptions });
}

/** Giá trị của preset đối chiếu với registry hiện hành; mục lỗi thời bị bỏ qua và được báo lại. */
export function applyPreset(registry: FormatRegistry, preset: Preset): { state: FormatState; ignored: string[] } {
  return reconcile(registry, preset.values, preset.exceptions);
}

/** "Lưu làm mặc định": giá trị đang dùng trở thành mặc định của chuẩn trên máy này. */
export function saveDefaults(store: KeyValueStore, standardId: StandardId, state: FormatState): void {
  store.set(defaultsKey(standardId), JSON.stringify({ values: state.values, exceptions: state.exceptions }));
}

export function clearDefaults(store: KeyValueStore, standardId: StandardId): void {
  store.remove(defaultsKey(standardId));
}

export function loadDefaults(store: KeyValueStore, registry: FormatRegistry, standardId: StandardId): { state: FormatState; ignored: string[] } {
  const saved = readJson(store, defaultsKey(standardId));
  if (typeof saved !== "object" || saved === null) {
    return { state: initialState(registry), ignored: [] };
  }
  const { values, exceptions } = saved as { values?: unknown; exceptions?: unknown };
  return reconcile(
    registry,
    typeof values === "object" && values !== null ? (values as Record<string, unknown>) : {},
    Array.isArray(exceptions) ? exceptions.filter((e): e is string => typeof e === "string") : [],
  );
}
