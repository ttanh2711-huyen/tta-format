import type { StandardProfile } from "../standards/types";
import type { KeyValueStore } from "./storage";

/** Danh sách cơ quan ban hành đã lưu, mỗi chuẩn một danh sách (T22 bổ sung quản lý, xuất và nhập). */

/** names: tên cơ quan ban hành; abbrs: chữ viết tắt cơ quan dùng trong ký hiệu văn bản. */
export type AgencyList = "names" | "abbrs" | "recipients";

const KEYS: Record<AgencyList, string> = { names: "nd30.agencies", abbrs: "nd30.agencyAbbrs", recipients: "nd30.recipients" };
export const agencyKey = (standardId: string, list: AgencyList): string => `${KEYS[list]}.${standardId}`;
const keyOf = (profile: StandardProfile, list: AgencyList): string => agencyKey(profile.id, list);

export function listAgencies(store: KeyValueStore, profile: StandardProfile, list: AgencyList = "names"): string[] {
  try {
    const parsed: unknown = JSON.parse(store.get(keyOf(profile, list)) ?? "[]");
    return Array.isArray(parsed) ? parsed.filter((item): item is string => typeof item === "string") : [];
  } catch {
    return [];
  }
}

export function saveAgency(store: KeyValueStore, profile: StandardProfile, name: string, list: AgencyList = "names"): { ok: true } | { ok: false; reason: string } {
  const clean = name.replace(/\s+/g, " ").trim();
  if (clean === "") return { ok: false, reason: "Hãy nhập tên cơ quan trước khi lưu." };
  const current = listAgencies(store, profile, list);
  if (current.some((item) => item.toLowerCase() === clean.toLowerCase())) return { ok: false, reason: `"${clean}" đã có trong danh sách.` };
  store.set(keyOf(profile, list), JSON.stringify([...current, clean]));
  return { ok: true };
}

export function removeAgency(store: KeyValueStore, profile: StandardProfile, name: string, list: AgencyList = "names"): void {
  store.set(keyOf(profile, list), JSON.stringify(listAgencies(store, profile, list).filter((item) => item !== name)));
}

/** Dòng nơi nhận gọn: bỏ gạch đầu dòng, dấu cuối dòng (; . ,) và khoảng trắng thừa. */
export function cleanRecipientLine(line: string): string {
  return line.replace(/\s+/g, " ").trim().replace(/^[-–—]\s*/, "").replace(/[;.,]+$/, "").trim();
}

/** Mục không phải tên nơi nhận: dòng "Lưu: ..." và "Như trên", "Như Điều ..." là câu trỏ lại, không đưa vào danh bạ. */
const NOT_A_NAME = /^(lưu\s*:|như\s)/i;

/**
 * Lưu các dòng nơi nhận vào danh bạ của chuẩn (T26): bỏ dòng trống, dòng "Lưu", "Như trên"; không lưu trùng (không phân biệt
 * hoa thường). Trả về số mục thêm mới.
 */
export function saveRecipients(store: KeyValueStore, profile: StandardProfile, lines: readonly string[]): { added: number } {
  const current = listAgencies(store, profile, "recipients");
  const next = [...current];
  for (const raw of lines) {
    const line = cleanRecipientLine(raw);
    if (line === "" || NOT_A_NAME.test(line) || next.some((item) => item.toLowerCase() === line.toLowerCase())) continue;
    next.push(line);
  }
  if (next.length > current.length) store.set(keyOf(profile, "recipients"), JSON.stringify(next));
  return { added: next.length - current.length };
}
