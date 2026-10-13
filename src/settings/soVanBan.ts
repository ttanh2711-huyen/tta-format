import type { StandardProfile } from "../standards/types";
import type { KeyValueStore } from "./storage";

/**
 * Sổ số văn bản (T26): ghi nhớ số văn bản đã dùng theo chuẩn, năm và loại (chữ viết tắt tên loại) để gợi ý số tiếp theo.
 * Số đánh liên tục từ 1 trong một năm cho từng loại; công văn (không có chữ viết tắt) dùng một bộ đếm chung. Dữ liệu chỉ
 * nằm trên máy này; muốn chuyển máy thì dùng xuất cấu hình.
 */

export interface Counter {
  year: number;
  typeKey: string;
  last: number;
}

const NO_TYPE = "-";

export function normalizeTypeKey(abbr: string | null | undefined): string {
  const clean = (abbr ?? "").replace(/\s+/g, "").toUpperCase();
  return clean === "" ? NO_TYPE : clean;
}

export const counterKey = (profile: StandardProfile, year: number, typeKey: string): string => `nd30.so.${profile.id}.${year}.${normalizeTypeKey(typeKey)}`;
const indexKey = (profile: StandardProfile): string => `nd30.so.index.${profile.id}`;

const isValidNumber = (n: number): boolean => Number.isInteger(n) && n >= 1;

export function lastNumber(store: KeyValueStore, profile: StandardProfile, year: number, abbr: string | null): number {
  const value = Number(store.get(counterKey(profile, year, normalizeTypeKey(abbr))));
  return isValidNumber(value) ? value : 0;
}

/** Số gợi ý cho văn bản kế tiếp. */
export function nextNumber(store: KeyValueStore, profile: StandardProfile, year: number, abbr: string | null): number {
  return lastNumber(store, profile, year, abbr) + 1;
}

function readIndex(store: KeyValueStore, profile: StandardProfile): string[] {
  try {
    const parsed: unknown = JSON.parse(store.get(indexKey(profile)) ?? "[]");
    return Array.isArray(parsed) ? parsed.filter((item): item is string => typeof item === "string") : [];
  } catch {
    return [];
  }
}

function write(store: KeyValueStore, profile: StandardProfile, year: number, typeKey: string, value: number): void {
  store.set(counterKey(profile, year, typeKey), String(value));
  const entry = `${year}|${normalizeTypeKey(typeKey)}`;
  const index = readIndex(store, profile);
  if (!index.includes(entry)) store.set(indexKey(profile), JSON.stringify([...index, entry]));
}

/** Ghi số vừa dùng: số tiếp theo chỉ tăng, không lùi khi gặp số nhỏ hơn (người dùng có thể tạo văn bản số cũ). */
export function recordNumber(store: KeyValueStore, profile: StandardProfile, year: number, abbr: string | null, used: number): void {
  if (!isValidNumber(used)) return;
  const typeKey = normalizeTypeKey(abbr);
  if (used > lastNumber(store, profile, year, typeKey)) write(store, profile, year, typeKey, used);
}

/** Sửa tay số cuối cùng đã dùng (kể cả giảm); false nếu số không hợp lệ. */
export function setLastNumber(store: KeyValueStore, profile: StandardProfile, year: number, abbr: string | null, last: number): boolean {
  if (!isValidNumber(last)) return false;
  write(store, profile, year, normalizeTypeKey(abbr), last);
  return true;
}

export function removeCounter(store: KeyValueStore, profile: StandardProfile, year: number, abbr: string | null): void {
  const typeKey = normalizeTypeKey(abbr);
  store.remove(counterKey(profile, year, typeKey));
  store.set(indexKey(profile), JSON.stringify(readIndex(store, profile).filter((entry) => entry !== `${year}|${typeKey}`)));
}

/** Các bộ đếm của chuẩn có số hợp lệ, năm mới nhất trước rồi theo tên loại. */
export function listCounters(store: KeyValueStore, profile: StandardProfile): Counter[] {
  const counters: Counter[] = [];
  for (const entry of readIndex(store, profile)) {
    const [yearText, typeKey] = entry.split("|");
    const year = Number(yearText);
    const last = lastNumber(store, profile, year, typeKey);
    if (Number.isInteger(year) && typeKey !== undefined && last > 0) counters.push({ year, typeKey, last });
  }
  return counters.sort((a, b) => b.year - a.year || a.typeKey.localeCompare(b.typeKey, "vi"));
}

/** Năm của văn bản: lấy từ ô năm nếu là số 4 chữ số, không thì năm hiện tại. */
export function yearOf(text: string, now: Date = new Date()): number {
  return /^\d{4}$/.test(text.trim()) ? Number(text.trim()) : now.getFullYear();
}
