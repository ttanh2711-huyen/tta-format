import { nextTabIndex } from "./tabsLogic";

export interface TabSpec {
  id: string;
  label: string;
  /** Mã vùng nội dung mà tab điều khiển; nhiều tab có thể dùng chung một vùng. */
  panelId: string;
  /** Biểu tượng kiểu ribbon: các nét vẽ (thuộc tính d của path SVG, khung 24×24). */
  icon?: readonly string[];
}

export interface TabsOptions {
  /** Tên của thanh tab cho trình đọc màn hình. */
  label?: string;
  /** Lớp CSS của thanh tab (mặc định "tabs"). */
  className?: string;
}

const SVG_NS = "http://www.w3.org/2000/svg";

function iconElement(paths: readonly string[]): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("class", "tab-icon");
  for (const d of paths) {
    const path = document.createElementNS(SVG_NS, "path");
    path.setAttribute("d", d);
    svg.append(path);
  }
  return svg;
}

export interface TabsHandle {
  /** Đặt tab đang chọn (không gọi `onSelect`). */
  setActive(id: string): void;
  active(): string;
}

/** Mã phần tử của nút tab, để vùng nội dung trỏ về bằng `aria-labelledby`. */
export const tabButtonId = (id: string): string => `tta-tab-${id}`;

/**
 * Thanh tab theo mẫu ARIA (T25): `tablist` với các `tab`, tab đang chọn là điểm dừng duy nhất của phím Tab; mũi tên,
 * Home, End di chuyển và chọn tab. Bấm hoặc dùng phím chỉ gọi `onSelect`; người gọi quyết định có đổi tab thật không
 * (ví dụ phải hỏi xác nhận trước) rồi gọi `setActive`.
 */
export function mountTabs(container: HTMLElement, tabs: readonly TabSpec[], initial: string, onSelect: (id: string) => void, options: TabsOptions = {}): TabsHandle {
  let active = tabs.some((tab) => tab.id === initial) ? initial : tabs[0].id;
  const list = document.createElement("div");
  list.className = options.className ?? "tabs";
  list.setAttribute("role", "tablist");
  list.setAttribute("aria-label", options.label ?? "Chức năng của TTA-Format");
  const buttons = tabs.map((tab) => {
    const button = document.createElement("button");
    button.type = "button";
    button.id = tabButtonId(tab.id);
    button.className = "tab";
    button.setAttribute("role", "tab");
    button.setAttribute("aria-controls", tab.panelId);
    if (tab.icon) button.append(iconElement(tab.icon));
    button.append(Object.assign(document.createElement("span"), { className: "tab-label", textContent: tab.label }));
    return button;
  });

  function paint(): void {
    buttons.forEach((button, index) => {
      const selected = tabs[index].id === active;
      button.setAttribute("aria-selected", String(selected));
      button.tabIndex = selected ? 0 : -1;
    });
  }

  buttons.forEach((button, index) => {
    button.addEventListener("click", () => onSelect(tabs[index].id));
    button.addEventListener("keydown", (event) => {
      const target = nextTabIndex(index, tabs.length, event.key);
      if (target === null) return;
      event.preventDefault();
      buttons[target].focus();
      onSelect(tabs[target].id);
    });
  });
  list.append(...buttons);
  container.replaceChildren(list);
  paint();

  return {
    setActive(id) {
      if (!tabs.some((tab) => tab.id === id)) return;
      active = id;
      paint();
    },
    active: () => active,
  };
}
