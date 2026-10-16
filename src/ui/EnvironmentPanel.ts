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
 * Danh sách "Môi trường Word" (T25): từng tính năng chạy được hay không trên bản Word đang mở, kèm lý do. Trả về dòng tóm tắt
 * và số tính năng không chạy được.
 */
export function mountEnvironmentPanel(container: HTMLElement, isSupported: IsSetSupported = officeIsSupported()): { unsupported: number; summary: string } {
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
  const summary = `Phiên bản ${appVersion()} · ${unsupported > 0 ? `${unsupported} tính năng không chạy được` : degraded > 0 ? `${degraded} tính năng chạy một phần` : "đủ tính năng"}`;
  container.replaceChildren(create("p", { className: "note", textContent: summary }), list);
  return { unsupported, summary };
}
