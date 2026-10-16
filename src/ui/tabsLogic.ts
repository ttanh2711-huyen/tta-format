import type { KeyValueStore } from "../settings/storage";

/**
 * Phần thuần tuý của thanh tab (T25), tách khỏi DOM để kiểm thử: phím di chuyển giữa các tab theo hướng dẫn
 * ARIA (mũi tên, Home, End, có vòng) và nhớ tab, mục con đang mở trên máy này.
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

const ACTIVE_TAB_KEY = "tta.activeTab";
const SECTION_PREFIX = "tta.section.";

/** Lựa chọn đã lưu ở khoá `key` nếu còn nằm trong `ids`, không thì lựa chọn đầu tiên. */
function loadChoice(store: KeyValueStore, key: string, ids: readonly string[]): string {
  const saved = store.get(key);
  return saved !== null && ids.includes(saved) ? saved : ids[0];
}

/** Tab đã chọn lần trước nếu còn nằm trong danh sách tab hiện có, không thì tab đầu tiên. */
export function loadActiveTab(store: KeyValueStore, ids: readonly string[]): string {
  return loadChoice(store, ACTIVE_TAB_KEY, ids);
}

export function saveActiveTab(store: KeyValueStore, id: string): void {
  store.set(ACTIVE_TAB_KEY, id);
}

/** Mục con đang mở của tab `tab` (mỗi tab nhớ riêng). */
export function loadActiveSection(store: KeyValueStore, tab: string, ids: readonly string[]): string {
  return loadChoice(store, `${SECTION_PREFIX}${tab}`, ids);
}

export function saveActiveSection(store: KeyValueStore, tab: string, id: string): void {
  store.set(`${SECTION_PREFIX}${tab}`, id);
}
