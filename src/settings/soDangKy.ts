import type { StandardProfile } from "../standards/types";
import type { KeyValueStore } from "./storage";

/**
 * Sổ đăng ký văn bản đi (T27): mỗi lần tạo văn bản thành công, add-in ghi một dòng (số, ký hiệu, ngày tháng, tên loại,
 * trích yếu, nơi nhận, người ký) theo chuẩn. Dữ liệu chỉ nằm trên máy này; xuất ra Excel hoặc bảng Word để lưu, in.
 */

export interface RegisterEntry {
  id: string;
  /** Năm của văn bản (theo ngày ghi trên văn bản, không có thì năm tạo). */
  year: number;
  /** Thời điểm ghi sổ (ISO). */
  createdAt: string;
  soKyHieu: string;
  noiNgay: string;
  loai: string;
  trichYeu: string;
  noiNhan: string;
  nguoiKy: string;
}

export type NewEntry = Omit<RegisterEntry, "id" | "createdAt">;

/** Giữ tối đa bấy nhiêu dòng mỗi chuẩn (bộ nhớ cục bộ của Word có hạn); quá thì bỏ dòng cũ nhất. */
export const MAX_ENTRIES = 3000;

export const REGISTER_HEADER = ["STT", "Số, ký hiệu", "Địa danh, ngày tháng", "Tên loại", "Trích yếu nội dung", "Nơi nhận", "Người ký", "Ngày ghi sổ"] as const;

export const registerKey = (profile: StandardProfile): string => `nd30.dk.${profile.id}`;
const clean = (text: string): string => text.replace(/\s+/g, " ").trim();

function isEntry(item: unknown): item is RegisterEntry {
  if (typeof item !== "object" || item === null) return false;
  const e = item as Record<string, unknown>;
  return typeof e.id === "string" && typeof e.year === "number" && typeof e.createdAt === "string" && ["soKyHieu", "noiNgay", "loai", "trichYeu", "noiNhan", "nguoiKy"].every((k) => typeof e[k] === "string");
}

/** Đọc các dòng hợp lệ từ một giá trị JSON bất kỳ (dùng cả khi nhập cấu hình). */
export function parseEntries(value: unknown): RegisterEntry[] {
  return Array.isArray(value) ? value.filter(isEntry) : [];
}

function readAll(store: KeyValueStore, profile: StandardProfile): RegisterEntry[] {
  try {
    return parseEntries(JSON.parse(store.get(registerKey(profile)) ?? "[]"));
  } catch {
    return [];
  }
}

function writeAll(store: KeyValueStore, profile: StandardProfile, entries: RegisterEntry[]): void {
  const ordered = [...entries].sort((a, b) => a.createdAt.localeCompare(b.createdAt));
  store.set(registerKey(profile), JSON.stringify(ordered.slice(Math.max(0, ordered.length - MAX_ENTRIES))));
}

export function addEntry(store: KeyValueStore, profile: StandardProfile, entry: NewEntry, now: Date = new Date()): RegisterEntry {
  const full: RegisterEntry = { ...entry, id: `${now.getTime().toString(36)}${Math.random().toString(36).slice(2, 7)}`, createdAt: now.toISOString() };
  writeAll(store, profile, [...readAll(store, profile), full]);
  return full;
}

export function removeEntry(store: KeyValueStore, profile: StandardProfile, id: string): void {
  writeAll(store, profile, readAll(store, profile).filter((entry) => entry.id !== id));
}

/** Gộp các dòng nhập từ cấu hình: dòng trùng mã giữ bản đang có. Trả số dòng mới. */
export function mergeEntries(store: KeyValueStore, profile: StandardProfile, incoming: readonly RegisterEntry[]): number {
  const current = readAll(store, profile);
  const ids = new Set(current.map((e) => e.id));
  const added = incoming.filter((e) => !ids.has(e.id));
  writeAll(store, profile, [...current, ...added]);
  return added.length;
}

/** Các dòng của sổ (lọc theo năm nếu có), mới nhất trước. */
export function listEntries(store: KeyValueStore, profile: StandardProfile, year?: number): RegisterEntry[] {
  return readAll(store, profile)
    .filter((entry) => year === undefined || entry.year === year)
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

export function registerYears(store: KeyValueStore, profile: StandardProfile): number[] {
  return [...new Set(readAll(store, profile).map((entry) => entry.year))].sort((a, b) => b - a);
}

/** Dòng sổ từ dữ liệu của văn bản vừa tạo. */
export function entryFromDocument(doc: { typeName: string; so: string; placeDate: string; subject: string; recipients: readonly string[]; signer: string }, now: Date = new Date()): NewEntry {
  const year = /năm\s+(\d{4})/.exec(doc.placeDate)?.[1];
  return {
    year: year ? Number(year) : now.getFullYear(),
    soKyHieu: clean(doc.so),
    noiNgay: clean(doc.placeDate),
    loai: clean(doc.typeName),
    trichYeu: clean(doc.subject),
    noiNhan: doc.recipients.map(clean).filter((line) => line !== "").join("; "),
    nguoiKy: clean(doc.signer),
  };
}

/** Bảng để xuất: dòng tiêu đề rồi các dòng theo thứ tự ghi sổ (cũ trước), có số thứ tự. */
export function registerRows(entries: readonly RegisterEntry[]): string[][] {
  const ordered = [...entries].sort((a, b) => a.createdAt.localeCompare(b.createdAt));
  return [
    [...REGISTER_HEADER],
    ...ordered.map((e, i) => [String(i + 1), e.soKyHieu, e.noiNgay, e.loai, e.trichYeu, e.noiNhan, e.nguoiKy, new Date(e.createdAt).toLocaleDateString("vi-VN")]),
  ];
}
