import { insertFooterBlock } from "../office/insertFooter";
import { isInsertHeaderSupported } from "../office/insertHeader";
import { buildPartyFooterBlock, partyFooterFromOptions, validatePartyFooter, type PartyFooterInput } from "../standards/hd05/footerBlock";
import { MAX_POSITION_LINES, buildFooterBlock, findAcademicTitle, footerInputFromOptions, validateFooterInput, type FooterBlockInput } from "../ooxml/footerBlock";
import type { FormatValues } from "../settings/formatOptions";
import type { AgencyProfile } from "../settings/profile";
import type { StandardProfile } from "../standards/types";

export interface FooterFormDeps {
  currentStandard: () => StandardProfile;
  values: () => FormatValues;
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

/** Form "Khối cuối văn bản" (T10): nơi nhận, quyền hạn, chức vụ, họ tên; chèn bảng ở cuối tài liệu. */
export type FooterCollect =
  | { ok: true; kind: "quocHieu"; input: FooterBlockInput; academicTitle: string | null }
  | { ok: true; kind: "dang"; input: PartyFooterInput; academicTitle: string | null }
  | { ok: false; error: string };

export function mountFooterForm(container: HTMLElement, deps: FooterFormDeps): { refresh: () => void; collect: () => FooterCollect; applyAgency: (agency: AgencyProfile) => void } {
  let recipients = "Như trên";
  let luu = "VT";
  let authority = "";
  let positionLines = [""];
  let signerName = "";
  let confirmTitle: string | null = null;
  let drafter = "";
  let withConfirmation = false;
  let confirmation = "";

  const body = create("div");
  const message = create("p", { id: "ff-message" });
  message.setAttribute("role", "status");
  container.append(create("h2", { textContent: "Khối cuối văn bản" }), body, message);

  function show(kind: "ok" | "warning" | "error" | "", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  /** Đọc và kiểm tra các ô đang nhập; dùng chung cho nút "Chèn khối cuối" và trình tạo công văn. */
  function collect(): FooterCollect {
    const profile = deps.currentStandard();
    const academicTitle = findAcademicTitle(signerName);
    if (profile.headerKind === "dang") {
      const party = partyFooterFromOptions(profile, deps.values(), {
        recipients: recipients.split(/\r?\n/),
        luu,
        drafter,
        authority,
        positionLines,
        signerName,
        confirmation: withConfirmation ? confirmation : "",
      });
      const partyProblems = validatePartyFooter(party);
      return partyProblems.length > 0 ? { ok: false, error: partyProblems.join(" ") } : { ok: true, kind: "dang", input: party, academicTitle };
    }
    const input = footerInputFromOptions(profile, deps.values(), { recipients: recipients.split(/\r?\n/), luu, authority, positionLines, signerName });
    const problems = validateFooterInput(input);
    return problems.length > 0 ? { ok: false, error: problems.join(" ") } : { ok: true, kind: "quocHieu", input, academicTitle };
  }

  async function insert(force: boolean): Promise<void> {
    const collected = collect();
    if (!collected.ok) {
      show("error", collected.error);
      return;
    }
    const xml = collected.kind === "dang" ? buildPartyFooterBlock(collected.input) : buildFooterBlock(collected.input);
    if (!force && collected.academicTitle) {
      confirmTitle = collected.academicTitle;
      render();
      return;
    }
    confirmTitle = null;
    try {
      const result = await insertFooterBlock(xml, signerName.trim());
      render();
      show(result.verified ? "ok" : "warning", result.verified ? "Đã chèn khối cuối văn bản ở cuối tài liệu." : "Đã chèn nhưng đọc lại không thấy họ tên người ký trong bảng cuối; hãy kiểm tra tài liệu.");
    } catch (error) {
      show("error", `Không chèn được khối cuối: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  function text(id: string, placeholder: string, value: string, onInput: (v: string) => void): HTMLInputElement {
    const input = create("input", { type: "text", id, value, placeholder });
    input.addEventListener("input", () => onInput(input.value));
    return input;
  }

  function authoritySection(profile: StandardProfile): HTMLElement {
    const section = create("fieldset", {}, create("legend", { textContent: "Quyền hạn ký" }));
    const choices = [{ abbr: "", meaning: "ký trực tiếp" }, ...profile.signAuthority];
    for (const choice of choices) {
      const radio = create("input", { type: "radio", name: "ff-authority", id: `ff-authority-${choice.abbr || "none"}`, value: choice.abbr, checked: authority === choice.abbr });
      radio.addEventListener("change", () => {
        authority = choice.abbr;
      });
      const label = create("label", { title: choice.meaning }, radio, ` ${choice.abbr || "Không"}`);
      label.style.marginRight = "8px";
      section.append(label);
    }
    return section;
  }

  function render(): void {
    const profile = deps.currentStandard();
    body.replaceChildren();
    const party = profile.headerKind === "dang";
    const list = create("textarea", { id: "ff-recipients", rows: 3, value: recipients, placeholder: "Nơi nhận, mỗi nơi một dòng" });
    list.addEventListener("input", () => {
      recipients = list.value;
    });
    body.append(
      create("p", { className: "note", textContent: party ? "Chèn bảng nơi nhận (gạch chân) và chữ ký (chức vụ không đậm) ở cuối tài liệu. Add-in tự thêm gạch đầu dòng, dấu ; và dấu . cuối mục." : "Chèn bảng nơi nhận và chữ ký ở cuối tài liệu. Add-in tự thêm gạch đầu dòng, dấu ; và dấu . cuối mục." }),
      create("div", {}, list),
      create("div", {}, "Lưu: ", text("ff-luu", "VT, ...", luu, (v) => (luu = v))),
      authoritySection(profile),
    );
    positionLines.forEach((line, index) => {
      body.append(create("div", {}, text(`ff-position-${index}`, `Chức vụ, dòng ${index + 1}`, line, (v) => (positionLines[index] = v))));
    });
    body.append(
      button("Thêm dòng chức vụ", () => {
        if (positionLines.length < MAX_POSITION_LINES) positionLines = [...positionLines, ""];
        render();
      }, { id: "ff-add-position", disabled: positionLines.length >= MAX_POSITION_LINES }),
      button("Bớt dòng chức vụ", () => {
        if (positionLines.length > 1) positionLines = positionLines.slice(0, -1);
        render();
      }, { id: "ff-remove-position", disabled: positionLines.length <= 1 }),
      create("div", {}, text("ff-name", "Họ và tên người ký (không ghi học hàm, học vị)", signerName, (v) => (signerName = v))),
    );
    if (party) {
      const toggle = create("input", { type: "checkbox", id: "ff-with-confirm", checked: withConfirmation });
      toggle.addEventListener("change", () => {
        withConfirmation = toggle.checked;
        render();
      });
      body.append(
        create("div", {}, text("ff-drafter", "Ký hiệu người soạn, số bản phát hành (nếu có)", drafter, (v) => (drafter = v))),
        create("div", {}, create("label", {}, toggle, " Có dòng xác nhận chữ ký")),
      );
      if (withConfirmation) body.append(create("div", {}, text("ff-confirmation", "Họ tên, chức vụ người được xác nhận", confirmation, (v) => (confirmation = v))));
    }
    if (confirmTitle !== null) {
      const box = create("div", { className: "ff-box", id: "ff-confirm" });
      box.setAttribute("role", "alert");
      box.append(
        create("p", { textContent: `Họ tên có "${confirmTitle}". Theo quy định, chữ ký chỉ ghi họ và tên, không ghi học hàm, học vị.` }),
        button("Vẫn chèn", () => void insert(true)),
        button("Sửa lại", () => {
          confirmTitle = null;
          render();
        }),
      );
      body.append(box);
    }
    const insertButton = button("Chèn khối cuối", () => void insert(false), { id: "ff-insert" });
    if (!isInsertHeaderSupported()) {
      insertButton.disabled = true;
      show("error", "Chèn khối cuối cần Microsoft 365 bản có WordApi 1.3 trở lên.");
    }
    body.append(insertButton);
  }

  render();
  return {
    refresh: () => {
      show("", "");
      render();
    },
    collect,
    applyAgency: (agency) => {
      positionLines = agency.positionLines.length > 0 ? [...agency.positionLines] : [""];
      if (agency.signerName !== "") signerName = agency.signerName;
      show("", "");
      render();
    },
  };
}
