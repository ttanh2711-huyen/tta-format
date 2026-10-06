/** Kho khoá-giá trị đơn giản cho cài đặt cục bộ; tách khỏi trình duyệt để test được. */
export interface KeyValueStore {
  get(key: string): string | null;
  set(key: string, value: string): void;
  remove(key: string): void;
}

export function memoryStore(): KeyValueStore {
  const data = new Map<string, string>();
  return {
    get: (key) => data.get(key) ?? null,
    set: (key, value) => {
      data.set(key, value);
    },
    remove: (key) => {
      data.delete(key);
    },
  };
}

/**
 * Kho lưu trong localStorage của task pane. localStorage có thể không dùng được (chế độ riêng tư,
 * chặn dữ liệu trang): khi đó rơi về bộ nhớ tạm, cài đặt chỉ còn trong phiên làm việc.
 */
export function browserStore(): KeyValueStore {
  const fallback = memoryStore();
  return {
    get(key) {
      try {
        return window.localStorage.getItem(key) ?? fallback.get(key);
      } catch {
        return fallback.get(key);
      }
    },
    set(key, value) {
      try {
        window.localStorage.setItem(key, value);
      } catch {
        fallback.set(key, value);
      }
    },
    remove(key) {
      fallback.remove(key);
      try {
        window.localStorage.removeItem(key);
      } catch {
        // không có localStorage thì chỉ cần xoá ở bộ nhớ tạm
      }
    },
  };
}
