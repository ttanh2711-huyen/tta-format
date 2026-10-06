import { formatBody, isFormatBodySupported, type FormatScope, type LocalFormatting } from "../office/formatBody";
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

function radioGroup<T extends string>(
  legend: string,
  name: string,
  choices: ReadonlyArray<{ value: T; label: string }>,
): { group: HTMLFieldSetElement; value: () => T } {
  const radios = choices.map((choice, index) =>
    create("input", { type: "radio", name, id: `${name}-${choice.value}`, value: choice.value, checked: index === 0 }),
  );
  const group = create("fieldset", { id: `${name}-group` }, create("legend", { textContent: legend }));
  choices.forEach((choice, index) => group.append(create("label", {}, radios[index], ` ${choice.label}`)));
  return { group, value: () => (radios.find((radio) => radio.checked)?.value ?? choices[0].value) as T };
}

/** Bảng "Định dạng nội dung" (T6): áp style nội dung của chuẩn đang chọn lên văn bản có sẵn. */
export function mountFormatPanel(container: HTMLElement, currentStandard: () => StandardProfile): void {
  const scope = radioGroup<FormatScope>("Phạm vi", "fb-scope", [
    { value: "selection", label: "Phần đang chọn" },
    { value: "document", label: "Cả tài liệu" },
  ]);
  const local = radioGroup<LocalFormatting>("Định dạng cục bộ có sẵn (đậm, nghiêng, căn lề…)", "fb-local", [
    { value: "keep", label: "Giữ lại" },
    { value: "clear", label: "Xoá để theo đúng style" },
  ]);
  const apply = create("button", { id: "fb-apply", type: "button", textContent: "Định dạng nội dung" });
  const message = create("p", { id: "fb-message" });
  message.setAttribute("role", "status");

  container.append(
    create("h2", { textContent: "Định dạng nội dung" }),
    create("p", {
      className: "note",
      textContent:
        "Áp style nội dung của chuẩn đang chọn lên các đoạn thường. Bỏ qua đoạn trống, đoạn trong bảng, tiêu đề có sẵn, danh sách tự động và đoạn đã dùng style của bộ chuẩn. Cần cài bộ style trước.",
    }),
    scope.group,
    local.group,
    apply,
    message,
  );

  function show(kind: "ok" | "warning" | "error", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  if (!isFormatBodySupported()) {
    apply.disabled = true;
    show("error", "Định dạng nội dung cần Microsoft 365 bản 2308 trở lên (WordApi 1.6). Hãy cập nhật Word rồi mở lại.");
    return;
  }

  async function onApply(): Promise<void> {
    apply.disabled = true;
    show("ok", "Đang định dạng…");
    try {
      const result = await formatBody(currentStandard(), { scope: scope.value(), localFormatting: local.value() });
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
