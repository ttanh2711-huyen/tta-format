import { HD05 } from "./hd05/profile";
import { ND30 } from "./nd30/profile";
import type { StandardId, StandardProfile } from "./types";

export type { StandardId, StandardProfile } from "./types";
export { HD05, ND30 };

/** Các chuẩn văn bản được hỗ trợ, theo thứ tự hiển thị. */
export const STANDARDS: readonly StandardProfile[] = [ND30, HD05];

export function getStandard(id: StandardId): StandardProfile {
  const found = STANDARDS.find((standard) => standard.id === id);
  if (!found) {
    throw new Error(`Không có chuẩn văn bản "${id}".`);
  }
  return found;
}
