import { iconElement } from "./Tabs";

export interface QuickAction {
  id: string;
  label: string;
  /** Một dòng mô tả ngắn dưới nhãn. */
  hint: string;
  icon: readonly string[];
  /** Chạy ngay (ô màu nhấn) hay chỉ mở mục tương ứng. */
  runs: boolean;
  onClick: () => void;
}

/**
 * Tab "Nhanh": các việc hay làm nhất, mỗi việc một lần bấm. Việc "chạy ngay" mở đúng mục rồi bấm nút ở đó để kết quả
 * hiện tại chỗ; việc "mở nhanh" chỉ đưa tới mục cần điền.
 */
export function mountQuickActions(container: HTMLElement, groups: ReadonlyArray<{ title: string; actions: readonly QuickAction[] }>): void {
  for (const group of groups) {
    const grid = document.createElement("div");
    grid.className = "qa-grid";
    for (const action of group.actions) {
      const button = document.createElement("button");
      button.type = "button";
      button.id = `qa-${action.id}`;
      button.className = action.runs ? "qa-tile qa-run" : "qa-tile";
      const label = Object.assign(document.createElement("span"), { className: "qa-label", textContent: action.label });
      const hint = Object.assign(document.createElement("span"), { className: "qa-hint", textContent: action.hint });
      button.append(iconElement(action.icon, "qa-icon"), label, hint);
      button.addEventListener("click", action.onClick);
      grid.append(button);
    }
    container.append(Object.assign(document.createElement("h3"), { className: "qa-title", textContent: group.title }), grid);
  }
}
