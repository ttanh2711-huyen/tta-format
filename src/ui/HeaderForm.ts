import { insertHeaderBlock, isInsertHeaderSupported, startsWithTable } from "../office/insertHeader";
import { QUOC_HIEU, buildHeaderBlock, headerInputFromOptions, MAX_PARENT_LINES, validateHeaderInput } from "../ooxml/headerBlock";
import { listAgencies, removeAgency, saveAgency } from "../settings/agencies";
import type { FormatValues } from "../settings/formatOptions";
import type { KeyValueStore } from "../settings/storage";
import type { StandardProfile } from "../standards/types";

export interface HeaderFormDeps {
  currentStandard: () => StandardProfile;
  values: () => FormatValues;
  store: KeyValueStore;
}

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

function button(text: string, onClick: () => void, props: Partial<HTMLButtonElement> = {}): HTMLButtonElement {
  const element = create("button", { type: "button", textContent: text, ...props });
  element.addEventListener("click", onClick);
  return element;
}

/** Form "Khối đầu văn bản" (T7): cơ quan chủ quản 0 đến 3 dòng, cơ quan ban hành, chèn bảng quốc hiệu và tiêu ngữ. */
export function mountHeaderForm(container: HTMLElement, deps: HeaderFormDeps): { refresh: () => void } {
  let hasParent = false;
  let parentLines = [""];
  let issuer = "";
  let confirmInsert = false;

  const body = create("div");
  const message = create("p", { id: "hf-message" });
  message.setAttribute("role", "status");
  container.append(create("h2", { textContent: "Khối đầu văn bản" }), body, message);

  function show(kind: "ok" | "warning" | "error" | "", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  async function insert(force: boolean): Promise<void> {
    const profile = deps.currentStandard();
    const input = headerInputFromOptions(profile, deps.values(), { parentLines: hasParent ? parentLines : [], issuer });
    const problems = validateHeaderInput(input);
    if (problems.length > 0) {
      show("error", problems.join(" "));
      return;
    }
    try {
      if (!force && (await startsWithTable())) {
        confirmInsert = true;
        render();
        return;
      }
      confirmInsert = false;
      const result = await insertHeaderBlock(buildHeaderBlock(input), QUOC_HIEU);
      render();
      show(result.verified ? "ok" : "warning", result.verified ? "Đã chèn khối đầu văn bản ở đầu tài liệu." : "Đã chèn nhưng đọc lại không thấy quốc hiệu trong bảng đầu tiên; hãy kiểm tra tài liệu.");
    } catch (error) {
      show("error", `Không chèn được khối đầu: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  function parentSection(): HTMLElement {
    const toggle = create("input", { type: "checkbox", id: "hf-has-parent", checked: hasParent });
    toggle.addEventListener("change", () => {
      hasParent = toggle.checked;
      render();
    });
    const section = create("div", {}, create("label", {}, toggle, " Có cơ quan chủ quản"));
    if (!hasParent) return section;
    parentLines.forEach((line, index) => {
      const input = create("input", { type: "text", id: `hf-parent-${index}`, value: line, placeholder: `Cơ quan chủ quản, dòng ${index + 1}` });
      input.addEventListener("input", () => {
        parentLines[index] = input.value;
      });
      section.append(create("div", {}, input));
    });
    section.append(
      button("Thêm dòng", () => {
        if (parentLines.length < MAX_PARENT_LINES) parentLines = [...parentLines, ""];
        render();
      }, { id: "hf-add-line", disabled: parentLines.length >= MAX_PARENT_LINES }),
      button("Bớt dòng", () => {
        if (parentLines.length > 1) parentLines = parentLines.slice(0, -1);
        render();
      }, { id: "hf-remove-line", disabled: parentLines.length <= 1 }),
    );
    return section;
  }

  function issuerSection(profile: StandardProfile): HTMLElement {
    const input = create("input", { type: "text", id: "hf-issuer", value: issuer, placeholder: "Cơ quan ban hành" });
    input.addEventListener("input", () => {
      issuer = input.value;
    });
    const section = create("div", {}, create("div", { textContent: "Cơ quan ban hành" }), input, " ");
    section.append(
      button("Lưu vào danh sách", () => {
        const saved = saveAgency(deps.store, profile, issuer);
        if (saved.ok) render();
        show(saved.ok ? "ok" : "error", saved.ok ? "Đã lưu vào danh sách cơ quan." : saved.reason);
      }, { id: "hf-save-agency" }),
    );
    const agencies = listAgencies(deps.store, profile);
    if (agencies.length > 0) {
      const list = create("div", { className: "hf-agencies" });
      for (const name of agencies) {
        list.append(
          button(name, () => {
            issuer = name;
            render();
          }),
          button("×", () => {
            removeAgency(deps.store, profile, name);
            render();
          }, { title: `Xoá "${name}" khỏi danh sách`, className: "hf-remove" }),
        );
      }
      section.append(list);
    }
    return section;
  }

  function render(): void {
    const profile = deps.currentStandard();
    body.replaceChildren();
    if (profile.headerKind !== "quocHieu") {
      body.append(create("p", { className: "note", textContent: `Khối đầu của chuẩn ${profile.name} (tiêu đề Đảng, cơ quan cấp trên) sẽ có ở bước TD1.` }));
      return;
    }
    body.append(
      create("p", { className: "note", textContent: "Chèn bảng cơ quan, quốc hiệu và tiêu ngữ ở đầu tài liệu. Cỡ chữ, độ dài đường kẻ và lề lấy từ Tuỳ chọn nhanh." }),
      parentSection(),
      issuerSection(profile),
    );
    if (confirmInsert) {
      const box = create("div", { className: "hf-box", id: "hf-confirm" });
      box.setAttribute("role", "alert");
      box.append(
        create("p", { textContent: "Tài liệu đang bắt đầu bằng một bảng, có thể đã có khối đầu. Chèn thêm sẽ cho hai khối đầu." }),
        button("Vẫn chèn", () => void insert(true)),
        button("Huỷ", () => {
          confirmInsert = false;
          render();
        }),
      );
      body.append(box);
    }
    const insertButton = button("Chèn khối đầu", () => void insert(false), { id: "hf-insert" });
    if (!isInsertHeaderSupported()) {
      insertButton.disabled = true;
      show("error", "Chèn khối đầu cần Microsoft 365 bản có WordApi 1.3 trở lên.");
    }
    body.append(insertButton);
  }

  render();
  return {
    refresh: () => {
      show("", "");
      render();
    },
  };
}
