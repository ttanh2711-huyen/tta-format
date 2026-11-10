import { iconElement } from "./Tabs";

export interface QuickAction {
  id: string;
  label: string;
  /** Mô tả đầy đủ, hiện khi rê chuột (không chiếm chỗ trên giao diện). */
  hint: string;
  /** Biểu tượng nét vẽ; chỉ ô ("tiles") mới có. */
  icon?: readonly string[];
  onClick: () => void;
}

export interface QuickGroup {
  title: string;
  /** tiles: ô lớn cho việc chạy ngay, hai cột; chips: viên nhỏ xếp hàng để mở mục, tiết kiệm chỗ. */
  kind: "tiles" | "chips";
  actions: readonly QuickAction[];
}

/**
 * Tab "Nhanh": các việc hay làm nhất, mỗi việc một lần bấm. Ô "chạy ngay" đi tới đúng mục rồi bấm nút ở đó để kết quả
 * hiện tại chỗ; viên "mở mục" chỉ đưa tới mục cần điền.
 */
export function mountQuickActions(container: HTMLElement, groups: readonly QuickGroup[]): void {
  for (const group of groups) {
    const list = document.createElement("div");
    list.className = group.kind === "tiles" ? "qa-grid" : "qa-chips";
    for (const action of group.actions) {
      const button = document.createElement("button");
      button.type = "button";
      button.id = `qa-${action.id}`;
      button.className = group.kind === "tiles" ? "qa-tile" : "qa-chip";
      button.title = action.hint;
      if (group.kind === "tiles" && action.icon) button.append(iconElement(action.icon, "qa-icon"));
      button.append(Object.assign(document.createElement("span"), { className: "qa-label", textContent: action.label }));
      button.addEventListener("click", action.onClick);
      list.append(button);
    }
    container.append(Object.assign(document.createElement("h3"), { className: "qa-title", textContent: group.title }), list);
  }
}
