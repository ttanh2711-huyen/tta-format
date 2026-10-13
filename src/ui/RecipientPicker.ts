import { cleanRecipientLine, listAgencies, removeAgency, saveRecipients } from "../settings/agencies";
import type { KeyValueStore } from "../settings/storage";
import type { StandardProfile } from "../standards/types";

function create<K extends keyof HTMLElementTagNameMap>(tag: K, props: Partial<HTMLElementTagNameMap[K]> = {}, ...children: Array<Node | string>): HTMLElementTagNameMap[K] {
  const element = document.createElement(tag);
  Object.assign(element, props);
  element.append(...children);
  return element;
}

/**
 * Danh bạ nơi nhận (T26): dưới ô nhập nơi nhận có các nút tên đã lưu; bấm tên thì thêm một dòng vào ô (không thêm trùng),
 * "Lưu các dòng vào danh bạ" ghi các dòng đang nhập để lần sau chọn lại. Danh bạ riêng cho từng chuẩn, nằm trong
 * cấu hình xuất và nhập được.
 */
export function recipientPicker(store: KeyValueStore, profile: StandardProfile, textarea: HTMLTextAreaElement): HTMLElement {
  const box = create("div", { className: "recipient-picker" });
  const note = create("p", { className: "note" });

  function addLine(name: string): void {
    const lines = textarea.value.split(/\r?\n/).map(cleanRecipientLine);
    if (lines.some((line) => line.toLowerCase() === name.toLowerCase())) return;
    const base = textarea.value.replace(/\s+$/, "");
    textarea.value = base === "" ? name : `${base}\n${name}`;
    // Cho ô nhập biết giá trị đã đổi (các biểu mẫu giữ trạng thái qua sự kiện input).
    textarea.dispatchEvent(new Event("input", { bubbles: true }));
  }

  function render(): void {
    const names = listAgencies(store, profile, "recipients");
    const chips = create("div", { className: "hf-agencies" });
    for (const name of names) {
      const pick = create("button", { type: "button", textContent: name, title: `Thêm "${name}" vào nơi nhận` });
      pick.addEventListener("click", () => addLine(name));
      const drop = create("button", { type: "button", textContent: "×", className: "hf-remove", title: `Xoá "${name}" khỏi danh bạ` });
      drop.addEventListener("click", () => {
        removeAgency(store, profile, name, "recipients");
        render();
      });
      chips.append(pick, drop);
    }
    const save = create("button", { type: "button", textContent: "Lưu các dòng vào danh bạ" });
    save.addEventListener("click", () => {
      const { added } = saveRecipients(store, profile, textarea.value.split(/\r?\n/));
      render();
      note.textContent = added > 0 ? `Đã lưu ${added} nơi nhận mới vào danh bạ.` : "Không có nơi nhận mới để lưu.";
    });
    box.replaceChildren(
      names.length > 0 ? create("div", { className: "note", textContent: "Danh bạ nơi nhận (bấm tên để thêm):" }) : create("div", { className: "note", textContent: "Danh bạ nơi nhận còn trống; gõ các nơi nhận rồi lưu để chọn lại lần sau." }),
      chips,
      save,
      note,
    );
  }

  render();
  return box;
}
