import type { StandardProfile } from "../standards/types";
import type { KeyValueStore } from "./storage";

/** Danh sách cơ quan ban hành đã lưu, mỗi chuẩn một danh sách (T22 bổ sung quản lý, xuất và nhập). */

/** names: tên cơ quan ban hành; abbrs: chữ viết tắt cơ quan dùng trong ký hiệu văn bản. */
export type AgencyList = "names" | "abbrs";

export const agencyKey = (standardId: string, list: AgencyList): string => (list === "names" ? `nd30.agencies.${standardId}` : `nd30.agencyAbbrs.${standardId}`);
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
