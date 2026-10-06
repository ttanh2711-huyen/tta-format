import { STANDARDS } from "../standards";
import type { StandardId, StandardProfile } from "../standards/types";
import type { KeyValueStore } from "./storage";

const KEY = "nd30.defaultStandard";

/** Chuẩn dùng cho tài liệu mới: lựa chọn gần nhất của người dùng trên máy này. */
export function loadDefaultStandard(store: KeyValueStore): StandardProfile {
  const saved = store.get(KEY);
  return STANDARDS.find((standard) => standard.id === saved) ?? STANDARDS[0];
}

export function saveDefaultStandard(store: KeyValueStore, id: StandardId): void {
  store.set(KEY, id);
}
