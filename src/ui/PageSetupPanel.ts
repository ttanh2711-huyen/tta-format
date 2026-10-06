import { applyPageSetup, isPageSetupSupported } from "../office/pageSetup";
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
  const standardSelect = create("select", { id: "ps-standard" });
  for (const standard of STANDARDS) {
    standardSelect.append(create("option", { value: standard.id, textContent: standard.name }));
  }

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

  container.append(
    create("h2", { textContent: "Thiết lập trang chuẩn" }),
    create("label", {}, "Chuẩn văn bản ", standardSelect),
    ...rows,
    create("label", {}, duplex, " In hai mặt (lề đối xứng) ", duplexNote),
    create("label", {}, font, " Đặt phông mặc định Times New Roman"),
    apply,
    message,
  );

  const currentStandard = (): StandardProfile => STANDARDS.find((s) => s.id === standardSelect.value) ?? STANDARDS[0];

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

  standardSelect.addEventListener("change", fillDefaults);
  fillDefaults();

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
