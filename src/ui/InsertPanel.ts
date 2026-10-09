import { appendAppendix, insertAtSelection, isInsertExtrasSupported } from "../office/insertExtras";
import { appendixNumberingXml, buildAppendix } from "../ooxml/appendix";
import { buildUrgencyMark } from "../ooxml/urgencyMark";
import type { FormatValues } from "../settings/formatOptions";
import type { StandardProfile } from "../standards/types";

export interface InsertPanelDeps {
  currentStandard: () => StandardProfile;
  values: () => FormatValues;
}

function create<K extends keyof HTMLElementTagNameMap>(tag: K, props: Partial<HTMLElementTagNameMap[K]> = {}, ...children: Array<Node | string>): HTMLElementTagNameMap[K] {
  const element = document.createElement(tag);
  Object.assign(element, props);
  element.append(...children);
  return element;
}

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/**
 * Bảng "Chèn dấu và phụ lục" (T21): chèn khung dấu chỉ mức độ khẩn tại con trỏ và chèn trang phụ lục ở cuối tài liệu.
 * Phụ lục chỉ hiện với chuẩn có style phụ lục.
 */
export function mountInsertPanel(container: HTMLElement, deps: InsertPanelDeps): { refresh: () => void } {
  let markIndex = 0;
  const appendix = { title: "", reference: "" };
  const body = create("div");
  const message = create("p", { id: "ip-message" });
  message.setAttribute("role", "status");
  container.append(create("h2", { textContent: "Chèn dấu khẩn và phụ lục" }), body, message);

  function show(kind: "ok" | "warning" | "error" | "", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  async function run(action: () => Promise<{ verified: boolean; note: string }>): Promise<void> {
    try {
      const result = await action();
      show(result.verified ? "ok" : "warning", result.note);
    } catch (error) {
      show("error", `Không chèn được: ${errorText(error)}`);
    }
  }

  function render(): void {
    const profile = deps.currentStandard();
    body.replaceChildren();
    const marks = profile.urgencyMarks;
    const sizeRange = profile.elements.dauKhan?.sizePt;
    const select = create("select", { id: "ip-mark" });
    marks.forEach((mark, index) => select.append(create("option", { value: String(index), textContent: mark.widthMm ? `${mark.text} (${mark.widthMm} × ${mark.heightMm} mm)` : mark.text })));
    select.value = String(Math.min(markIndex, marks.length - 1));
    select.addEventListener("change", () => {
      markIndex = Number(select.value);
    });
    const insertMark = create("button", { type: "button", id: "ip-insert-mark", textContent: "Chèn khung tại con trỏ" });
    insertMark.addEventListener("click", () => {
      const mark = marks[Number(select.value)];
      if (!mark) return;
      const size = typeof deps.values().fontSize === "number" ? (deps.values().fontSize as number) : (sizeRange?.max ?? 14);
      const clamped = sizeRange ? Math.min(sizeRange.max, Math.max(sizeRange.min, size)) : size;
      void run(() => insertAtSelection(buildUrgencyMark({ text: mark.text, widthMm: mark.widthMm, heightMm: mark.heightMm, sizePt: clamped }), mark.text));
    });
    body.append(
      create("p", { className: "note", textContent: "Đặt con trỏ dưới số và ký hiệu (công văn: dưới trích yếu) rồi chèn khung chữ nhật viền đơn, chữ in hoa đậm." }),
      create("div", {}, select, " ", insertMark),
    );
    if (!profile.elements.phuLucSo) {
      body.append(create("p", { className: "note", textContent: `Trang phụ lục của chuẩn ${profile.name} chưa có trong add-in.` }));
      return;
    }
    const prefix = profile.styles.prefix;
    const title = create("input", { type: "text", id: "ip-appendix-title", value: appendix.title, placeholder: "Tên phụ lục, ví dụ Danh sách đại biểu" });
    title.addEventListener("input", () => (appendix.title = title.value));
    const reference = create("input", { type: "text", id: "ip-appendix-ref", value: appendix.reference, placeholder: "Chỉ dẫn, ví dụ Kèm theo Công văn số ... ngày ..." });
    reference.addEventListener("input", () => (appendix.reference = reference.value));
    const insertAppendix = create("button", { type: "button", id: "ip-insert-appendix", textContent: "Chèn phụ lục ở cuối tài liệu" });
    insertAppendix.addEventListener("click", () => {
      void run(() =>
        appendAppendix(`${prefix}PhuLucSo`, `${prefix}PhuLucTieuDe`, appendix.title.replace(/\s+/g, " ").trim().toUpperCase(), (number) => ({
          bodyXml: buildAppendix({ title: appendix.title, reference: appendix.reference, number, stylePrefix: prefix }),
          numberingXml: appendixNumberingXml(number),
        })),
      );
    });
    body.append(
      create("p", { className: "note", textContent: "Phụ lục bắt đầu trang mới, nhãn \"Phụ lục I, II...\" tự đánh số nối tiếp các phụ lục đã có." }),
      create("div", {}, title),
      create("div", {}, reference),
      insertAppendix,
    );
  }

  if (!isInsertExtrasSupported()) {
    body.append(create("p", { className: "error", textContent: "Chèn dấu và phụ lục cần Microsoft 365 bản có WordApi 1.3 trở lên." }));
  } else {
    render();
  }
  return { refresh: () => (isInsertExtrasSupported() ? (show("", ""), render()) : undefined) };
}
