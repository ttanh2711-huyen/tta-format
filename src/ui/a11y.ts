/**
 * Trợ năng (T23): mọi ô nhập, ô chọn, vùng văn bản phải có tên để trình đọc màn hình đọc được. Giao diện dựng động nên
 * một bộ quan sát gắn `aria-label` cho điều khiển chưa có tên, lấy từ chữ gợi ý (placeholder) hoặc tiêu đề (title).
 */

/** Phần của điều khiển mà quy tắc đặt tên cần biết (tách khỏi DOM để kiểm thử). */
export interface ControlInfo {
  type?: string;
  placeholder?: string;
  title?: string;
  ariaLabel?: string | null;
  ariaLabelledBy?: string | null;
  /** Số nhãn `<label>` gắn với điều khiển. */
  labelCount: number;
}

/** Tên cần gắn cho điều khiển, hoặc null nếu đã có tên hay không có gì để lấy. */
export function accessibleNameFor(control: ControlInfo): string | null {
  if (control.type === "hidden") return null;
  if (control.labelCount > 0 || (control.ariaLabel ?? "").trim() !== "" || (control.ariaLabelledBy ?? "").trim() !== "") return null;
  const name = (control.placeholder ?? "").trim() || (control.title ?? "").trim();
  return name === "" ? null : name;
}

const SELECTOR = "input, select, textarea";

/** Gắn tên cho các điều khiển chưa có tên trong `root`; trả về số điều khiển đã gắn. */
export function labelControls(root: ParentNode): number {
  let changed = 0;
  root.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>(SELECTOR).forEach((control) => {
    const name = accessibleNameFor({
      type: control instanceof HTMLInputElement ? control.type : undefined,
      placeholder: "placeholder" in control ? control.placeholder : undefined,
      title: control.title,
      ariaLabel: control.getAttribute("aria-label"),
      ariaLabelledBy: control.getAttribute("aria-labelledby"),
      labelCount: control.labels?.length ?? 0,
    });
    if (name !== null) {
      control.setAttribute("aria-label", name);
      changed += 1;
    }
  });
  return changed;
}

/** Gắn tên ngay và mỗi khi giao diện đổi (các bảng dựng lại nội dung khi đổi chuẩn hoặc bấm nút). */
export function watchControls(root: HTMLElement): void {
  labelControls(root);
  new MutationObserver(() => {
    labelControls(root);
  }).observe(root, { childList: true, subtree: true });
}
