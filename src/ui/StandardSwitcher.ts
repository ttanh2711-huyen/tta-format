import type { DocumentSnapshot } from "../office/documentStandard";
import { loadDefaultStandard, saveDefaultStandard } from "../settings/defaultStandard";
import type { KeyValueStore } from "../settings/storage";
import { detectStandard, resolveStandard, type UnsureReason } from "../standards/detectStandard";
import { STANDARDS } from "../standards";
import type { StandardId, StandardProfile } from "../standards/types";

export interface StandardSwitcherDeps {
  store: KeyValueStore;
  /** Đọc tài liệu đang mở; bỏ trống nếu Word không hỗ trợ thuộc tính tuỳ chỉnh. */
  readSnapshot?: () => Promise<DocumentSnapshot>;
  writeStandard?: (id: StandardId) => Promise<void>;
  writeNoAsk?: () => Promise<void>;
}

export interface StandardSwitcherHandle {
  currentStandard: () => StandardProfile;
  onChange: (callback: () => void) => void;
  /** Đọc tài liệu, nhận diện chuẩn hoặc hỏi lại. Gọi một lần sau khi dựng các bảng khác. */
  start: () => Promise<void>;
  /** Người dùng chọn chuẩn (bấm tab chuẩn): đổi ngay, hoặc hỏi xác nhận khi tài liệu đã có nội dung theo chuẩn khác. */
  choose: (id: StandardId) => void;
}

const REASON_TEXT: Record<UnsureReason, string> = {
  noTitle: "Không thấy tiêu đề chuẩn (quốc hiệu hoặc tiêu đề Đảng) ở đầu văn bản.",
  bothTitles: "Đầu văn bản có cả hai loại tiêu đề.",
  similarTitle: "Tiêu đề ở đầu văn bản gần giống tiêu đề của một chuẩn nhưng không khớp hẳn (có thể sai chính tả hoặc ngắt dòng lạ).",
};

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

function button(text: string, onClick: () => void): HTMLButtonElement {
  const element = create("button", { type: "button", textContent: text });
  element.addEventListener("click", onClick);
  return element;
}

/**
 * Quản lý "chuẩn văn bản" đang dùng (T25: không còn ô chọn riêng, người dùng chọn bằng tab). Chuẩn chỉ đổi khi người
 * dùng chọn hoặc khi tiêu đề ở đầu văn bản rõ ràng; không chắc thì hỏi. Đổi chuẩn không tự sửa nội dung văn bản.
 */
export function mountStandardSwitcher(container: HTMLElement, deps: StandardSwitcherDeps): StandardSwitcherHandle {
  let current = loadDefaultStandard(deps.store);
  /** Chuẩn mà tài liệu đang theo (đã ghi hoặc nhận diện được); null nếu chưa rõ. */
  let documentStandard: StandardId | null = null;
  let hasContent = false;
  const callbacks: Array<() => void> = [];

  const dialog = create("div", { id: "ss-dialog" });
  const message = create("p", { id: "ss-message" });
  message.setAttribute("role", "status");
  container.append(dialog, message);

  function show(kind: "ok" | "warning" | "error" | "", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  function select(profile: StandardProfile): void {
    current = profile;
    for (const callback of callbacks) callback();
  }

  async function commit(profile: StandardProfile): Promise<void> {
    select(profile);
    saveDefaultStandard(deps.store, profile.id);
    documentStandard = profile.id;
    dialog.replaceChildren();
    const base = `Đã chuyển sang chuẩn ${profile.name}. Nội dung văn bản không bị sửa; chỉ bộ tuỳ chọn và các nút theo chuẩn mới.`;
    if (!deps.writeStandard) {
      show("ok", base);
      return;
    }
    try {
      await deps.writeStandard(profile.id);
      show("ok", `${base} Đã ghi chuẩn vào tài liệu.`);
    } catch (error) {
      show("warning", `${base} Chưa ghi được chuẩn vào tài liệu: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  function confirmSwitch(target: StandardProfile): void {
    const from = STANDARDS.find((standard) => standard.id === documentStandard);
    const box = create("div", { className: "ss-box", id: "ss-confirm" });
    box.setAttribute("role", "alert");
    box.append(
      create("p", {
        textContent: `Tài liệu này đã có nội dung theo chuẩn ${from?.name ?? "khác"}. Đổi sang ${target.name} chỉ đổi các tuỳ chọn và nút của add-in; nội dung hiện có không tự đổi, nên có thể không còn đúng thể thức của chuẩn mới.`,
      }),
      button("Đổi chuẩn", () => {
        void commit(target);
      }),
      button("Giữ nguyên", () => {
        dialog.replaceChildren();
        show("", "");
      }),
    );
    dialog.replaceChildren(box);
  }

  function choose(id: StandardId): void {
    const target = STANDARDS.find((standard) => standard.id === id);
    if (!target || target.id === current.id) return;
    if (hasContent && documentStandard !== null && documentStandard !== target.id) {
      confirmSwitch(target);
    } else {
      void commit(target);
    }
  }

  function ask(reason: UnsureReason, suggestion: StandardId | undefined): void {
    const noAsk = create("input", { type: "checkbox", id: "ss-no-ask" });
    const box = create("div", { className: "ss-box", id: "ss-ask" });
    box.setAttribute("role", "alert");
    box.append(create("strong", { textContent: "Văn bản này theo chuẩn nào?" }), create("p", { textContent: REASON_TEXT[reason] }));
    for (const standard of STANDARDS) {
      const label = standard.id === suggestion ? `${standard.name} (gợi ý)` : standard.name;
      box.append(
        button(label, () => {
          void commit(standard).then(() => (noAsk.checked ? deps.writeNoAsk?.() : undefined));
        }),
      );
    }
    box.append(
      button("Để sau", () => {
        dialog.replaceChildren();
        if (noAsk.checked) void deps.writeNoAsk?.();
      }),
      create("label", {}, noAsk, " Không hỏi lại cho tài liệu này"),
    );
    dialog.replaceChildren(box);
    show("warning", "Chưa xác định được chuẩn của tài liệu; hãy chọn một chuẩn. Add-in không tự áp dụng định dạng khi chưa chắc.");
  }

  async function start(): Promise<void> {
    if (!deps.readSnapshot) return;
    let snapshot: DocumentSnapshot;
    try {
      snapshot = await deps.readSnapshot();
    } catch (error) {
      show("warning", `Không đọc được tài liệu để nhận diện chuẩn: ${error instanceof Error ? error.message : String(error)}`);
      return;
    }
    hasContent = snapshot.hasContent;
    const resolution = resolveStandard({
      stored: snapshot.stored,
      noAsk: snapshot.noAsk,
      detection: detectStandard(snapshot.headerTexts),
      fallbackId: current.id,
    });
    if (resolution.action === "ask") {
      ask(resolution.reason, resolution.suggestion);
      return;
    }
    const profile = STANDARDS.find((standard) => standard.id === resolution.standardId) ?? current;
    if (profile.id !== current.id) select(profile);
    if (resolution.source !== "default") documentStandard = profile.id;
    if (resolution.source === "document") show("ok", `Tài liệu đã ghi chuẩn ${profile.name}.`);
    if (resolution.source === "detected") show("ok", `Nhận diện từ tiêu đề đầu văn bản: chuẩn ${profile.name}.`);
  }

  return { currentStandard: () => current, onChange: (callback) => void callbacks.push(callback), start, choose };
}
