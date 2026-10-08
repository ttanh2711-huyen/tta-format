import type { StandardProfile } from "../standards/types";
import type { KeyValueStore } from "./storage";

/** Danh sách cơ quan ban hành đã lưu, mỗi chuẩn một danh sách (T22 bổ sung quản lý, xuất và nhập). */

const keyOf = (profile: StandardProfile): string => `nd30.agencies.${profile.id}`;

export function listAgencies(store: KeyValueStore, profile: StandardProfile): string[] {
  try {
    const parsed: unknown = JSON.parse(store.get(keyOf(profile)) ?? "[]");
    return Array.isArray(parsed) ? parsed.filter((item): item is string => typeof item === "string") : [];
  } catch {
    return [];
  }
}

export function saveAgency(store: KeyValueStore, profile: StandardProfile, name: string): { ok: true } | { ok: false; reason: string } {
  const clean = name.replace(/\s+/g, " ").trim();
  if (clean === "") return { ok: false, reason: "Hãy nhập tên cơ quan trước khi lưu." };
  const current = listAgencies(store, profile);
  if (current.some((item) => item.toLowerCase() === clean.toLowerCase())) return { ok: false, reason: `"${clean}" đã có trong danh sách.` };
  store.set(keyOf(profile), JSON.stringify([...current, clean]));
  return { ok: true };
}

export function removeAgency(store: KeyValueStore, profile: StandardProfile, name: string): void {
  store.set(keyOf(profile), JSON.stringify(listAgencies(store, profile).filter((item) => item !== name)));
}
