import { formatBody, isFormatBodySupported, type FormatBodyOptions } from "../office/formatBody";
import type { StyleSpec } from "../ooxml/styleSpecs";
import type { StandardProfile } from "../standards/types";

function create<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  props: Partial<HTMLElementTagNameMap[K]> = {},
  ...children: Array<Node | string>
): HTMLElementTagNameMap[K] {
  const element = document.createElement(tag);
  Object.assign(element, props);
  element.append(...children);
  return element;
}

/** Bảng "Định dạng nội dung" (T6): áp style nội dung của chuẩn đang chọn lên văn bản có sẵn. */
export interface FormatPanelDeps {
  currentStandard: () => StandardProfile;
  /** Phạm vi và cách xử lý định dạng cục bộ, lấy từ nhóm "Hành vi" của tuỳ chọn nhanh. */
  behavior: () => Pick<FormatBodyOptions, "scope" | "localFormatting">;
  styleSpecs: () => StyleSpec[];
}

export function mountFormatPanel(container: HTMLElement, deps: FormatPanelDeps): void {
  const apply = create("button", { id: "fb-apply", type: "button", textContent: "Định dạng nội dung" });
  const message = create("p", { id: "fb-message" });
  message.setAttribute("role", "status");

  container.append(
    create("h2", { textContent: "Định dạng nội dung" }),
    create("p", {
      className: "note",
      textContent:
        "Áp style nội dung của chuẩn đang chọn lên các đoạn thường. Bỏ qua đoạn trống, đoạn trong bảng, tiêu đề có sẵn, danh sách tự động và đoạn đã dùng style của bộ chuẩn. Cần cài bộ style trước. Phạm vi và việc giữ hay xoá định dạng cục bộ chọn ở Tuỳ chọn nhanh, nhóm Hành vi.",
    }),
    apply,
    message,
  );

  function show(kind: "ok" | "warning" | "error", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  if (!isFormatBodySupported()) {
    apply.disabled = true;
    show("error", "Định dạng nội dung cần Word 2016 trở lên. Hãy cập nhật Word rồi mở lại.");
    return;
  }

  async function onApply(): Promise<void> {
    apply.disabled = true;
    show("ok", "Đang định dạng…");
    try {
      const result = await formatBody(deps.currentStandard(), { ...deps.behavior(), specs: deps.styleSpecs() });
      show(result.summary.kind, result.summary.text);
    } catch (error) {
      show("error", `Không định dạng được: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      apply.disabled = false;
    }
  }

  apply.addEventListener("click", () => {
    void onApply();
  });
}
