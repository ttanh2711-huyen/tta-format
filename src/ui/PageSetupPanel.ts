import { installStyles, isStyleInstallSupported } from "../office/installStyles";
import { applyPageSetup, isPageSetupSupported } from "../office/pageSetup";
import { summarizeApply } from "../office/pageSetupPlan";
import { summarizeInstall } from "../office/styleVerify";
import type { StyleSpec } from "../ooxml/styleSpecs";
import { toPageSetupSettings, type FormatValues } from "../settings/formatOptions";
import type { StandardProfile } from "../standards/types";

export interface PageSetupDeps {
  currentStandard: () => StandardProfile;
  /** Giá trị tuỳ chọn nhanh hiện hành (lề, khoảng cách header/footer, in hai mặt). */
  values: () => FormatValues;
  /** Bộ style theo tuỳ chọn nhanh hiện hành. */
  styleSpecs: () => StyleSpec[];
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

/**
 * Hai nút "Thiết lập trang" (T4) và "Cài bộ style" (T5). Chuẩn chọn ở công tắc đầu task pane; lề và các
 * giá trị khác lấy từ Tuỳ chọn nhanh nên chỉ có một nơi để chỉnh.
 */
export function mountPageSetupPanel(container: HTMLElement, deps: PageSetupDeps): void {
  const font = create("input", { id: "ps-font", type: "checkbox", checked: true });
  const apply = create("button", { id: "ps-apply", type: "button", textContent: "Thiết lập trang" });
  const message = create("p", { id: "ps-message" });
  message.setAttribute("role", "status");
  const styleButton = create("button", { id: "ps-install-styles", type: "button", textContent: "Cài bộ style" });
  const styleMessage = create("p", { id: "ps-style-message" });
  styleMessage.setAttribute("role", "status");

  container.append(
    create("h2", { textContent: "Thiết lập trang chuẩn" }),
    create("p", { className: "note", textContent: "Khổ A4 dọc cho mọi phần. Lề, khoảng cách header/footer và in hai mặt lấy từ Tuỳ chọn nhanh, nhóm Trang." }),
    create("label", {}, font, " Đặt phông mặc định theo chuẩn"),
    apply,
    message,
    create("h2", { textContent: "Bộ style" }),
    create("p", { className: "note", textContent: "Cài hoặc cập nhật các style của chuẩn đang chọn (ND30_*, HD05_*) theo Tuỳ chọn nhanh. Chạy lại thì cập nhật tại chỗ, không tạo bản sao; style Normal và Heading không bị đổi." }),
    styleButton,
    styleMessage,
  );

  function show(kind: "ok" | "warning" | "error" | "", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  function showStyle(kind: "ok" | "warning" | "error", text: string): void {
    styleMessage.className = kind;
    styleMessage.textContent = text;
  }

  async function onInstallStyles(): Promise<void> {
    styleButton.disabled = true;
    try {
      const result = await installStyles(deps.styleSpecs());
      const summary = summarizeInstall(result.items);
      showStyle(summary.kind, [summary.text, ...result.notes].join(" "));
    } catch (error) {
      showStyle("error", `Không cài được style: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      styleButton.disabled = false;
    }
  }

  if (isStyleInstallSupported()) {
    styleButton.addEventListener("click", () => {
      void onInstallStyles();
    });
  } else {
    styleButton.disabled = true;
    showStyle("error", "Cài style cần Microsoft 365 bản 2308 trở lên (WordApi 1.6). Hãy cập nhật Word rồi mở lại.");
  }

  if (!isPageSetupSupported()) {
    apply.disabled = true;
    show("error", "Tính năng này cần Microsoft 365 bản 2507 trở lên (WordApiDesktop 1.3). Hãy cập nhật Word rồi mở lại.");
    return;
  }

  async function onApply(): Promise<void> {
    const standard = deps.currentStandard();
    const settings = toPageSetupSettings(deps.values());
    apply.disabled = true;
    try {
      const result = await applyPageSetup(settings, font.checked ? { defaultFontName: standard.font.name } : {});
      const summary = summarizeApply(settings, result.sections, result.mismatches);
      show(summary.kind, [summary.text, result.fontNote].filter(Boolean).join(" "));
    } catch (error) {
      show("error", `Không thiết lập được trang: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      apply.disabled = false;
    }
  }

  apply.addEventListener("click", () => {
    void onApply();
  });
}
