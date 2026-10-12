import type { KeyValueStore } from "../settings/storage";

/**
 * Phần thuần tuý của thanh tab và mục xổ ra (T25), tách khỏi DOM để kiểm thử: phím di chuyển giữa các tab theo hướng dẫn
 * ARIA (mũi tên, Home, End, có vòng) và nhớ trạng thái mở của từng mục xổ ra trên máy này.
 */

/** Tab nhận tiêu điểm khi bấm `key` ở tab số `current`; null nếu phím này không đổi tab. */
export function nextTabIndex(current: number, count: number, key: string): number | null {
  if (count <= 0) return null;
  switch (key) {
    case "ArrowRight":
    case "ArrowDown":
      return (current + 1) % count;
    case "ArrowLeft":
    case "ArrowUp":
      return (current - 1 + count) % count;
    case "Home":
      return 0;
    case "End":
      return count - 1;
    default:
      return null;
  }
}

const ACCORDION_PREFIX = "tta.open.";
const ACTIVE_TAB_KEY = "tta.activeTab";

/** Mục xổ ra `id` đang mở không; chưa có ghi nhớ thì theo `fallback`. */
export function loadAccordionOpen(store: KeyValueStore, id: string, fallback: boolean): boolean {
  const saved = store.get(`${ACCORDION_PREFIX}${id}`);
  return saved === "1" ? true : saved === "0" ? false : fallback;
}

export function saveAccordionOpen(store: KeyValueStore, id: string, open: boolean): void {
  store.set(`${ACCORDION_PREFIX}${id}`, open ? "1" : "0");
}

/** Tab đã chọn lần trước nếu còn nằm trong danh sách tab hiện có, không thì tab đầu tiên. */
export function loadActiveTab(store: KeyValueStore, ids: readonly string[]): string {
  const saved = store.get(ACTIVE_TAB_KEY);
  return saved !== null && ids.includes(saved) ? saved : ids[0];
}

export function saveActiveTab(store: KeyValueStore, id: string): void {
  store.set(ACTIVE_TAB_KEY, id);
}
