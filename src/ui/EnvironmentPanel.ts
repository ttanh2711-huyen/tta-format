import { describeDegraded, describeMissing, evaluateFeatures, officeIsSupported, type IsSetSupported } from "../office/capabilities";

declare const __APP_VERSION__: string | undefined;

/** Số phiên bản của bản đang chạy (do bước build ghi vào); bản chạy thử trong test không có. */
export function appVersion(): string {
  return typeof __APP_VERSION__ === "string" ? __APP_VERSION__ : "chạy thử";
}

function create<K extends keyof HTMLElementTagNameMap>(tag: K, props: Partial<HTMLElementTagNameMap[K]> = {}, ...children: Array<Node | string>): HTMLElementTagNameMap[K] {
  const element = document.createElement(tag);
  Object.assign(element, props);
  element.append(...children);
  return element;
}

/**
 * Mục "Môi trường Word" (T23): liệt kê từng tính năng chạy được hay không trên bản Word đang mở, kèm lý do. Tự mở ra khi có
 * tính năng bị tắt để người dùng không phải đoán vì sao một nút không chạy.
 */
export function mountEnvironmentPanel(container: HTMLElement, isSupported: IsSetSupported = officeIsSupported()): { unsupported: number } {
  const statuses = evaluateFeatures(isSupported);
  const unsupported = statuses.filter((status) => !status.supported).length;
  const degraded = statuses.filter((status) => describeDegraded(status) !== "").length;
  const list = create("ul", { id: "env-list" });
  for (const status of statuses) {
    const note = describeMissing(status) || describeDegraded(status);
    list.append(
      create("li", { className: status.supported ? (note ? "warning" : "ok") : "error" }, create("strong", { textContent: status.supported ? (note ? "Chạy một phần: " : "Chạy được: ") : "Không chạy được: " }), status.feature.label, note ? create("div", { className: "note", textContent: note }) : ""),
    );
  }
  const summary = `Phiên bản ${appVersion()} · ${unsupported > 0 ? `Môi trường Word: ${unsupported} tính năng không chạy được` : degraded > 0 ? `Môi trường Word: ${degraded} tính năng chạy một phần` : "Môi trường Word: đủ tính năng"}`;
  const details = create("details", { id: "env-details", open: unsupported > 0 }, create("summary", { textContent: summary }), list);
  container.replaceChildren(details);
  return { unsupported };
}
