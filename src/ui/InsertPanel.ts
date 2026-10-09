import { appendAppendix, appendToEnd, countAppendices, countAppendicesByLabel, insertAtSelection, insertTextAtSelection, isInsertExtrasSupported } from "../office/insertExtras";
import { appendixNumberingXml, buildAppendix } from "../ooxml/appendix";
import { buildUrgencyMark } from "../ooxml/urgencyMark";
import { buildCirculationMark, buildContactInfo, buildDraftNote, buildNoPublicNote, buildPartyAppendix, NO_PUBLIC_NOTE, partyAppendixNumberingXml } from "../standards/hd05/extras";
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

const NAME_UPPER = (text: string): string => text.replace(/\s+/g, " ").trim().toUpperCase();

/**
 * Bảng "Chèn dấu khẩn, phụ lục và thành phần bổ sung" (T21, TD8): khung dấu chỉ mức độ khẩn tại con trỏ, trang phụ lục ở
 * cuối tài liệu và, với chuẩn Đảng, chỉ dẫn phạm vi lưu hành, chỉ dẫn dự thảo, dòng không phổ biến và thông tin liên hệ.
 */
export function mountInsertPanel(container: HTMLElement, deps: InsertPanelDeps): { refresh: () => void } {
  let markIndex = 0;
  const state = { appendixTitle: "", appendixRef: "", circulation: "", draft: "Dự thảo lần 1", contact: "" };
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

  const button = (id: string, text: string, onClick: () => void): HTMLButtonElement => {
    const element = create("button", { type: "button", id, textContent: text });
    element.addEventListener("click", onClick);
    return element;
  };
  const input = (id: string, placeholder: string, value: string, onInput: (v: string) => void): HTMLInputElement => {
    const element = create("input", { type: "text", id, value, placeholder });
    element.addEventListener("input", () => onInput(element.value));
    return element;
  };

  function sizeOf(profile: StandardProfile, elementId: string, fallback: number): number {
    const range = profile.elements[elementId]?.sizePt;
    return range ? range.min : fallback;
  }

  function renderUrgency(profile: StandardProfile): void {
    const marks = profile.urgencyMarks;
    const sizeRange = profile.elements.dauKhan?.sizePt;
    const select = create("select", { id: "ip-mark" });
    marks.forEach((mark, index) => select.append(create("option", { value: String(index), textContent: mark.widthMm ? `${mark.text} (${mark.widthMm} × ${mark.heightMm} mm)` : mark.text })));
    select.value = String(Math.min(markIndex, marks.length - 1));
    select.addEventListener("change", () => {
      markIndex = Number(select.value);
    });
    const insertMark = button("ip-insert-mark", "Chèn khung tại con trỏ", () => {
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
  }

  function renderAppendix(profile: StandardProfile): void {
    const prefix = profile.styles.prefix;
    const styled = Boolean(profile.elements.phuLucSo);
    const insertAppendix = button("ip-insert-appendix", "Chèn phụ lục ở cuối tài liệu", () => {
      const title = state.appendixTitle;
      const reference = state.appendixRef;
      void run(() =>
        appendAppendix(
          styled
            ? {
                count: () => countAppendices(`${prefix}PhuLucSo`),
                matches: (p) => p.style === `${prefix}PhuLucTieuDe` && p.text.trim() === NAME_UPPER(title),
                build: (number) => ({ bodyXml: buildAppendix({ title, reference, number, stylePrefix: prefix }), numberingXml: appendixNumberingXml(number) }),
              }
            : {
                count: countAppendicesByLabel,
                matches: (p) => p.text.trim() === NAME_UPPER(title),
                build: (number) => ({ bodyXml: buildPartyAppendix({ title, reference, number, sizePt: typeof deps.values().fontSize === "number" ? (deps.values().fontSize as number) : 14 }), numberingXml: partyAppendixNumberingXml(number) }),
              },
        ),
      );
    });
    body.append(
      create("p", { className: "note", textContent: 'Phụ lục bắt đầu trang mới, nhãn "Phụ lục I, II..." tự đánh số nối tiếp các phụ lục đã có.' + (styled ? "" : " Đánh số trang riêng cho từng phụ lục: đặt tay trong Word (Chèn > Số trang > Định dạng > Bắt đầu từ 1).") }),
      create("div", {}, input("ip-appendix-title", "Tên phụ lục, ví dụ Danh sách đại biểu", state.appendixTitle, (v) => (state.appendixTitle = v))),
      create("div", {}, input("ip-appendix-ref", "Chỉ dẫn, ví dụ Kèm theo Công văn số ... ngày ...", state.appendixRef, (v) => (state.appendixRef = v))),
      insertAppendix,
    );
  }

  function renderPartyExtras(profile: StandardProfile): void {
    const circulationSize = sizeOf(profile, "phamViLuuHanh", 12);
    const draftSize = sizeOf(profile, "duThao", 14);
    const contactSize = sizeOf(profile, "thongTinLienHe", 10);
    const contact = create("textarea", { id: "ip-contact", rows: 3, value: state.contact, placeholder: "Thông tin liên hệ, mỗi dòng một mục (địa chỉ, điện thoại, thư điện tử)" });
    contact.addEventListener("input", () => (state.contact = contact.value));
    body.append(
      create("h3", { textContent: "Thành phần bổ sung" }),
      create("p", { className: "note", textContent: `Chỉ dẫn phạm vi lưu hành (khung ${circulationSize} pt, góc phải trang đầu), chỉ dẫn dự thảo (${draftSize} pt nghiêng, dưới số và ký hiệu), thông tin liên hệ (${contactSize} pt, trang cuối). Đặt con trỏ đúng chỗ trước khi chèn tại con trỏ.` }),
      create("div", {}, input("ip-circulation", "Phạm vi lưu hành, ví dụ XEM XONG TRẢ LẠI", state.circulation, (v) => (state.circulation = v)), " ", button("ip-insert-circulation", "Chèn khung tại con trỏ", () => void run(() => insertAtSelection(buildCirculationMark(state.circulation, circulationSize), NAME_UPPER(state.circulation))))),
      create("div", {}, input("ip-draft", "Chỉ dẫn dự thảo", state.draft, (v) => (state.draft = v)), " ", button("ip-insert-draft", "Chèn tại con trỏ", () => void run(() => insertTextAtSelection(buildDraftNote(state.draft, draftSize), state.draft.trim())))),
      create("div", {}, button("ip-insert-nopublic", "Chèn dòng không phổ biến ở cuối", () => void run(() => appendToEnd(buildNoPublicNote(circulationSize), NO_PUBLIC_NOTE)))),
      create("div", {}, contact),
      button("ip-insert-contact", "Chèn thông tin liên hệ ở cuối", () => {
        const lines = state.contact.split(/\r?\n/);
        void run(() => appendToEnd(buildContactInfo(lines, contactSize), lines.find((line) => line.trim() !== "")?.trim() ?? ""));
      }),
    );
  }

  function render(): void {
    const profile = deps.currentStandard();
    body.replaceChildren();
    renderUrgency(profile);
    renderAppendix(profile);
    if (profile.elements.thongTinLienHe) renderPartyExtras(profile);
  }

  if (!isInsertExtrasSupported()) {
    body.append(create("p", { className: "error", textContent: "Chèn dấu và phụ lục cần Microsoft 365 bản có WordApi 1.3 trở lên." }));
  } else {
    render();
  }
  return { refresh: () => (isInsertExtrasSupported() ? (show("", ""), render()) : undefined) };
}
