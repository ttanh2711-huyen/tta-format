import { installStyles, isStyleInstallSupported } from "../office/installStyles";
import { applyPageSetup, isPageSetupSupported } from "../office/pageSetup";
import { summarizeInstall } from "../office/styleVerify";
import { buildStyleSpecs } from "../ooxml/styleSpecs";
import { parseMm, planPageSetup, summarizeApply } from "../office/pageSetupPlan";
import { STANDARDS } from "../standards";
import type { MarginSide, StandardProfile } from "../standards/types";

const SIDES: ReadonlyArray<{ side: MarginSide; label: string }> = [
  { side: "top", label: "Lề trên (mm)" },
  { side: "bottom", label: "Lề dưới (mm)" },
  { side: "left", label: "Lề trái (mm)" },
  { side: "right", label: "Lề phải (mm)" },
];

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

/** Bảng "Thiết lập trang chuẩn" trong task pane (T4). Chọn chuẩn tạm thời ở đây; T6B sẽ chuyển lên đầu task pane. */
export function mountPageSetupPanel(container: HTMLElement): void {
  // Dùng nút chọn (radio) thay cho ô thả xuống: ô thả xuống trong khung task pane của Word hay không mở được.
  const standardRadios = STANDARDS.map((standard, index) =>
    create("input", { type: "radio", name: "ps-standard", id: `ps-standard-${standard.id}`, value: standard.id, checked: index === 0 }),
  );
  const standardGroup = create("fieldset", { id: "ps-standard-group" }, create("legend", { textContent: "Chuẩn văn bản" }));
  STANDARDS.forEach((standard, index) => {
    standardGroup.append(create("label", {}, standardRadios[index], ` ${standard.name}`));
  });

  const inputs = new Map<MarginSide, HTMLInputElement>();
  const rows = SIDES.map(({ side, label }) => {
    const input = create("input", { id: `ps-${side}`, type: "text", inputMode: "decimal" });
    inputs.set(side, input);
    return create("label", {}, `${label} `, input);
  });

  const duplex = create("input", { id: "ps-duplex", type: "checkbox" });
  const duplexNote = create("span", { className: "note" });
  const font = create("input", { id: "ps-font", type: "checkbox", checked: true });
  const apply = create("button", { id: "ps-apply", type: "button", textContent: "Thiết lập trang" });
  const message = create("p", { id: "ps-message" });
  message.setAttribute("role", "status");
  const styleButton = create("button", { id: "ps-install-styles", type: "button", textContent: "Cài bộ style" });
  const styleMessage = create("p", { id: "ps-style-message" });
  styleMessage.setAttribute("role", "status");

  container.append(
    create("h2", { textContent: "Thiết lập trang chuẩn" }),
    standardGroup,
    ...rows,
    create("label", {}, duplex, " In hai mặt (lề đối xứng) ", duplexNote),
    create("label", {}, font, " Đặt phông mặc định Times New Roman"),
    apply,
    message,
    create("h2", { textContent: "Bộ style" }),
    create("p", { className: "note", textContent: "Cài hoặc cập nhật các style của chuẩn đang chọn (ND30_*, HD05_*). Chạy lại thì cập nhật tại chỗ, không tạo bản sao; style Normal và Heading không bị đổi." }),
    styleButton,
    styleMessage,
  );

  const currentStandard = (): StandardProfile => {
    const checked = standardRadios.find((radio) => radio.checked);
    return STANDARDS.find((s) => s.id === checked?.value) ?? STANDARDS[0];
  };

  function show(kind: "ok" | "warning" | "error" | "", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  function fillDefaults(): void {
    const standard = currentStandard();
    for (const { side } of SIDES) {
      const input = inputs.get(side);
      if (input) {
        input.value = String(standard.pageDefaults.marginsMm[side]).replace(".", ",");
      }
    }
    const supportsDuplex = standard.page.duplexBackMarginsMm !== undefined;
    duplex.disabled = !supportsDuplex;
    duplex.checked = false;
    duplexNote.textContent = supportsDuplex ? "" : "(chuẩn này không quy định)";
    show("", "");
  }

  for (const radio of standardRadios) {
    radio.addEventListener("change", fillDefaults);
  }
  fillDefaults();

  function showStyle(kind: "ok" | "warning" | "error", text: string): void {
    styleMessage.className = kind;
    styleMessage.textContent = text;
  }

  async function onInstallStyles(): Promise<void> {
    styleButton.disabled = true;
    try {
      const result = await installStyles(buildStyleSpecs(currentStandard()));
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
    const standard = currentStandard();
    const marginsMm: Partial<Record<MarginSide, number>> = {};
    for (const { side } of SIDES) {
      marginsMm[side] = parseMm(inputs.get(side)?.value ?? "");
    }
    const plan = planPageSetup(standard, { marginsMm, duplex: duplex.checked });
    if (!plan.ok) {
      show("error", `Chưa thiết lập: ${plan.errors.join(" ")}`);
      return;
    }
    apply.disabled = true;
    try {
      const result = await applyPageSetup(plan.settings, font.checked ? { defaultFontName: standard.font.name } : {});
      const summary = summarizeApply(plan.settings, result.sections, result.mismatches);
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
