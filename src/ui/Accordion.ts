import type { KeyValueStore } from "../settings/storage";
import { loadAccordionOpen, saveAccordionOpen } from "./tabsLogic";

export interface AccordionSpec {
  /** Mã ổn định để nhớ trạng thái mở, ví dụ "nd30.khoi-dau". */
  id: string;
  title: string;
  /** Dòng mô tả ngắn dưới tiêu đề khi đóng lại. */
  hint?: string;
  defaultOpen?: boolean;
}

export interface AccordionHandle {
  element: HTMLDetailsElement;
  /** Phần thân: nơi gắn nội dung. */
  body: HTMLElement;
}

/**
 * Mục xổ ra (T25): `<details>` có tiêu đề, mô tả ngắn và phần thân; bấm tiêu đề để mở hoặc đóng, trạng thái mở được nhớ
 * trên máy này. Dùng phần tử có sẵn của trình duyệt nên mở/đóng bằng Enter, Space và đọc được bằng trình đọc màn hình.
 */
export function createAccordion(spec: AccordionSpec, store: KeyValueStore): AccordionHandle {
  const element = document.createElement("details");
  element.className = "acc";
  element.dataset.accId = spec.id;
  element.open = loadAccordionOpen(store, spec.id, spec.defaultOpen ?? false);
  const summary = document.createElement("summary");
  const title = document.createElement("span");
  title.className = "acc-title";
  title.textContent = spec.title;
  summary.append(title);
  if (spec.hint) {
    const hint = document.createElement("span");
    hint.className = "acc-hint";
    hint.textContent = spec.hint;
    summary.append(hint);
  }
  const body = document.createElement("div");
  body.className = "acc-body";
  element.append(summary, body);
  element.addEventListener("toggle", () => saveAccordionOpen(store, spec.id, element.open));
  return { element, body };
}
