import { formatBody, isFormatBodySupported, StyleMissingError, type FormatBodyOptions, type FormatScope } from "../office/formatBody";
import { installStyles } from "../office/installStyles";
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

/**
 * Bảng "Định dạng nội dung" (T6): áp style của chuẩn đang chọn lên văn bản có sẵn, cho phần đang tô đen hoặc cả văn bản.
 * Tài liệu chưa có bộ style thì tự cài trước.
 */
export interface FormatPanelDeps {
  currentStandard: () => StandardProfile;
  /** Cách xử lý định dạng cục bộ, lấy từ nhóm "Hành vi" của tuỳ chọn nhanh. */
  behavior: () => Pick<FormatBodyOptions, "localFormatting">;
  styleSpecs: () => StyleSpec[];
}

export function mountFormatPanel(container: HTMLElement, deps: FormatPanelDeps): void {
  const applySelection = create("button", { id: "fb-selection", type: "button", textContent: "Định dạng phần tô đen", className: "btn-primary", title: "Bôi đen các đoạn cần định dạng trong văn bản rồi bấm" });
  const applyDocument = create("button", { id: "fb-apply", type: "button", textContent: "Định dạng cả văn bản" });
  const buttons = [applySelection, applyDocument];
  const message = create("p", { id: "fb-message" });
  message.setAttribute("role", "status");

  container.append(
    create("h2", { textContent: "Định dạng nội dung" }),
    create("p", {
      className: "note",
      textContent:
        "Nhận ra chương, mục, điều, khoản, điểm, thân bài rồi áp style của chuẩn đang chọn. Phần tô đen: đưa phông, cỡ chữ, căn lề, thụt dòng, giãn đoạn, giãn dòng về đúng chuẩn, kể cả đoạn đã có style chuẩn; giữ chữ đậm, nghiêng trong đoạn. Cả văn bản: chỉ áp cho đoạn chưa dùng style chuẩn. Luôn bỏ qua đoạn trống, trong bảng, tiêu đề Heading và danh sách tự động. Chưa có bộ style thì add-in tự cài.",
    }),
    create("div", { className: "rp-actions" }, ...buttons),
    message,
  );

  function show(kind: "ok" | "warning" | "error", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  if (!isFormatBodySupported()) {
    for (const button of buttons) button.disabled = true;
    show("error", "Định dạng nội dung cần Word 2016 trở lên. Hãy cập nhật Word rồi mở lại.");
    return;
  }

  async function onApply(scope: FormatScope): Promise<void> {
    for (const button of buttons) button.disabled = true;
    show("ok", "Đang định dạng…");
    const options = { ...deps.behavior(), scope, specs: deps.styleSpecs() };
    try {
      let installed = "";
      const result = await formatBody(deps.currentStandard(), options).catch(async (error: unknown) => {
        if (!(error instanceof StyleMissingError)) throw error;
        show("ok", "Tài liệu chưa có bộ style của chuẩn; đang cài…");
        await installStyles(options.specs);
        installed = "Đã cài bộ style của chuẩn. ";
        return formatBody(deps.currentStandard(), options);
      });
      show(result.summary.kind, installed + result.summary.text);
    } catch (error) {
      show("error", `Không định dạng được: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      for (const button of buttons) button.disabled = false;
    }
  }

  applySelection.addEventListener("click", () => void onApply("selection"));
  applyDocument.addEventListener("click", () => void onApply("document"));
}
