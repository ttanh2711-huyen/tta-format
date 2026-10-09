import { insertHeaderBlock, isInsertHeaderSupported, startsWithTable } from "../office/insertHeader";
import { buildTitleBody, subjectLine, titleInputFromOptions, validateTitleInput, type TitleBlockInput } from "../ooxml/titleBlock";
import { QUOC_HIEU, buildHeaderBlock, headerInputFromOptions, MAX_PARENT_LINES, validateHeaderInput, type HeaderBlockInput } from "../ooxml/headerBlock";
import { buildPartyHeaderBlock, MAX_PARENT_LINES as MAX_PARTY_PARENT_LINES, partyHeaderFromOptions, validatePartyHeader, type PartyHeaderInput } from "../standards/hd05/headerBlock";
import { placeDateText, soText } from "../rules/validators";
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

/** Ô số để trống nghĩa là chưa có (null); chữ không phải số thì giữ NaN để bị từ chối. */
function toInt(text: string): number | null {
  const trimmed = text.trim();
  return trimmed === "" ? null : Number(trimmed);
}

function button(text: string, onClick: () => void, props: Partial<HTMLButtonElement> = {}): HTMLButtonElement {
  const element = create("button", { type: "button", textContent: text, ...props });
  element.addEventListener("click", onClick);
  return element;
}

/** Form "Khối đầu văn bản" (T7): cơ quan chủ quản 0 đến 3 dòng, cơ quan ban hành, chèn bảng quốc hiệu và tiêu ngữ. */
export type HeaderCollect =
  | { ok: true; kind: "quocHieu"; input: HeaderBlockInput; titleInput: TitleBlockInput | undefined }
  | { ok: true; kind: "dang"; input: PartyHeaderInput }
  | { ok: false; error: string };

export function mountHeaderForm(container: HTMLElement, deps: HeaderFormDeps): { refresh: () => void; collect: () => HeaderCollect } {
  let hasParent = false;
  let parentLines = [""];
  let issuer = "";
  let confirmInsert = false;
  let withNumberRow = true;
  let withTitle = true;
  let variant: "letter" | "named" = "letter";
  const title = { typeName: "", subject: "", recipients: "" };
  const number = { so: "", typeAbbr: "", agencyAbbrs: "", day: "", month: "", year: "" };
  const PLACE_KEY = "nd30.place";
  let place = deps.store.get(PLACE_KEY) ?? "";

  const body = create("div");
  const message = create("p", { id: "hf-message" });
  message.setAttribute("role", "status");
  container.append(create("h2", { textContent: "Khối đầu văn bản" }), body, message);

  function show(kind: "ok" | "warning" | "error" | "", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  /** Đọc và kiểm tra các ô đang nhập; dùng chung cho nút "Chèn khối đầu" và trình tạo công văn. */
  function collect(): HeaderCollect {
    const profile = deps.currentStandard();
    let numberRow: { so: string; placeDate: string } | undefined;
    if (withNumberRow) {
      const so = soText(profile, { so: toInt(number.so), typeAbbr: number.typeAbbr.trim() === "" ? null : number.typeAbbr.trim(), agencyAbbrs: number.agencyAbbrs.split("-").map((part) => part.trim()).filter((part) => part !== "") });
      const date = placeDateText({ place, day: toInt(number.day), month: toInt(number.month), year: toInt(number.year) });
      if (!so.ok || !date.ok) return { ok: false, error: [so.ok ? "" : so.error, date.ok ? "" : date.error].filter(Boolean).join(" ") };
      numberRow = { so: so.text, placeDate: date.text };
    }
    if (profile.headerKind === "dang") {
      const partyInput = partyHeaderFromOptions(profile, deps.values(), { parentLines: hasParent ? parentLines : [], issuerLines: issuer.split(/\r?\n/), numberRow });
      const partyProblems = validatePartyHeader(partyInput);
      if (partyProblems.length > 0) return { ok: false, error: partyProblems.join(" ") };
      if (numberRow) deps.store.set(PLACE_KEY, place.trim());
      return { ok: true, kind: "dang", input: partyInput };
    }
    let subjectRow: { text: string } | undefined;
    let titleInput: TitleBlockInput | undefined;
    if (withTitle) {
      titleInput = titleInputFromOptions(profile, deps.values(), {
        variant,
        typeName: title.typeName,
        subject: title.subject,
        recipients: title.recipients.split(/\r?\n/),
      });
      const titleProblems = validateTitleInput(titleInput);
      if (titleProblems.length > 0) return { ok: false, error: titleProblems.join(" ") };
      if (variant === "letter") subjectRow = { text: subjectLine(title.subject) };
    }
    const input = headerInputFromOptions(profile, deps.values(), { parentLines: hasParent ? parentLines : [], issuer, numberRow, subjectRow });
    const problems = validateHeaderInput(input);
    if (problems.length > 0) return { ok: false, error: problems.join(" ") };
    if (numberRow) deps.store.set(PLACE_KEY, place.trim());
    return { ok: true, kind: "quocHieu", input, titleInput };
  }

  async function insert(force: boolean): Promise<void> {
    const collected = collect();
    if (!collected.ok) {
      show("error", collected.error);
      return;
    }
    const profile = deps.currentStandard();
    const xml = collected.kind === "dang" ? buildPartyHeaderBlock(collected.input) : buildHeaderBlock(collected.input) + (collected.titleInput ? buildTitleBody(collected.titleInput) : "");
    const expected = collected.kind === "dang" ? profile.identifyingTitle : QUOC_HIEU;
    try {
      if (!force && (await startsWithTable())) {
        confirmInsert = true;
        render();
        return;
      }
      confirmInsert = false;
      const result = await insertHeaderBlock(xml, expected);
      render();
      show(result.verified ? "ok" : "warning", result.verified ? "Đã chèn khối đầu văn bản ở đầu tài liệu." : "Đã chèn nhưng đọc lại không thấy tiêu đề trong bảng đầu tiên; hãy kiểm tra tài liệu.");
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
    const party = deps.currentStandard().headerKind === "dang";
    const maxLines = party ? MAX_PARTY_PARENT_LINES : MAX_PARENT_LINES;
    const noun = party ? "cơ quan cấp trên" : "cơ quan chủ quản";
    const section = create("div", {}, create("label", {}, toggle, ` Có ${noun}`));
    if (!hasParent) return section;
    parentLines.slice(0, maxLines).forEach((line, index) => {
      const input = create("input", { type: "text", id: `hf-parent-${index}`, value: line, placeholder: `${noun[0].toUpperCase()}${noun.slice(1)}, dòng ${index + 1}` });
      input.addEventListener("input", () => {
        parentLines[index] = input.value;
      });
      section.append(create("div", {}, input));
    });
    section.append(
      button("Thêm dòng", () => {
        if (parentLines.length < maxLines) parentLines = [...parentLines, ""];
        render();
      }, { id: "hf-add-line", disabled: parentLines.length >= maxLines }),
      button("Bớt dòng", () => {
        if (parentLines.length > 1) parentLines = parentLines.slice(0, -1);
        render();
      }, { id: "hf-remove-line", disabled: parentLines.length <= 1 }),
    );
    return section;
  }

  function issuerSection(profile: StandardProfile): HTMLElement {
    const input: HTMLInputElement | HTMLTextAreaElement =
      profile.headerKind === "dang"
        ? create("textarea", { id: "hf-issuer", rows: 2, value: issuer, placeholder: "Cơ quan ban hành (liên cơ quan: mỗi cơ quan một dòng)" })
        : create("input", { type: "text", id: "hf-issuer", value: issuer, placeholder: "Cơ quan ban hành" });
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

  /** Gợi ý ký hiệu: chữ viết tắt loại từ bảng loại văn bản, chữ viết tắt cơ quan đã lưu và bảng tra của chuẩn. */
  function suggestionSection(): HTMLElement {
    const profile = deps.currentStandard();
    const section = create("div", { className: "hf-suggest" });
    const types = [...new Set(profile.docTypes.map((type) => type.abbr).filter((abbr): abbr is string => abbr !== null))];
    if (types.length > 0) {
      section.append(create("div", { className: "note", textContent: "Gợi ý viết tắt loại:" }));
      for (const abbr of types) {
        section.append(
          button(abbr, () => {
            number.typeAbbr = abbr;
            render();
          }, { title: profile.docTypes.filter((type) => type.abbr === abbr).map((type) => type.name).join(", ") }),
        );
      }
    }
    const append = (abbr: string) => {
      number.agencyAbbrs = number.agencyAbbrs.trim() === "" ? abbr : `${number.agencyAbbrs.trim()}-${abbr}`;
      render();
    };
    const saved = listAgencies(deps.store, profile, "abbrs");
    section.append(
      create("div", {},
        button("Lưu viết tắt cơ quan", () => {
          const result = saveAgency(deps.store, profile, number.agencyAbbrs, "abbrs");
          if (result.ok) render();
          show(result.ok ? "ok" : "error", result.ok ? "Đã lưu chữ viết tắt cơ quan." : result.reason);
        }, { id: "hf-save-abbr" }),
      ),
    );
    if (saved.length > 0) {
      section.append(create("div", { className: "note", textContent: "Đã lưu:" }));
      for (const abbr of saved) {
        section.append(
          button(abbr, () => append(abbr)),
          button("×", () => {
            removeAgency(deps.store, profile, abbr, "abbrs");
            render();
          }, { title: `Xoá "${abbr}" khỏi danh sách`, className: "hf-remove" }),
        );
      }
    }
    if (profile.agencyCodes) {
      const details = create("details", {}, create("summary", { textContent: "Bảng ký hiệu cơ quan của chuẩn" }));
      for (const agency of profile.agencyCodes) details.append(button(agency.code, () => append(agency.code), { title: agency.name }));
      section.append(details);
    }
    return section;
  }

  function numberSection(): HTMLElement {
    const toggle = create("input", { type: "checkbox", id: "hf-with-number", checked: withNumberRow });
    toggle.addEventListener("change", () => {
      withNumberRow = toggle.checked;
      render();
    });
    const section = create("div", {}, create("label", {}, toggle, " Thêm hàng số, ký hiệu và ngày tháng"));
    if (!withNumberRow) return section;
    const field = (id: string, label: string, value: string, onInput: (v: string) => void, width = 60): HTMLElement => {
      const input = create("input", { type: "text", id, value, inputMode: "text" });
      input.style.width = `${width}px`;
      input.addEventListener("input", () => onInput(input.value));
      return create("label", { className: "hf-field" }, `${label} `, input);
    };
    section.append(
      create("div", { className: "note", textContent: "Để trống số hoặc ngày nếu chưa vào sổ, chưa ký; add-in chừa chỗ trống." }),
      field("hf-so", "Số", number.so, (v) => (number.so = v), 50),
      field("hf-type", "Viết tắt loại", number.typeAbbr, (v) => (number.typeAbbr = v), 60),
      field("hf-agency", "Viết tắt cơ quan (nối bằng -)", number.agencyAbbrs, (v) => (number.agencyAbbrs = v), 110),
      suggestionSection(),
      create("div", {}, field("hf-place", "Địa danh", place, (v) => (place = v), 130)),
      create("div", {},
        field("hf-day", "Ngày", number.day, (v) => (number.day = v), 36),
        field("hf-month", "tháng", number.month, (v) => (number.month = v), 36),
        field("hf-year", "năm", number.year, (v) => (number.year = v), 50),
        button("Hôm nay", () => {
          const now = new Date();
          number.day = String(now.getDate());
          number.month = String(now.getMonth() + 1);
          number.year = String(now.getFullYear());
          render();
        }, { id: "hf-today" }),
      ),
    );
    return section;
  }

  function titleSection(): HTMLElement {
    const toggle = create("input", { type: "checkbox", id: "hf-with-title", checked: withTitle });
    toggle.addEventListener("change", () => {
      withTitle = toggle.checked;
      render();
    });
    const section = create("div", {}, create("label", {}, toggle, " Thêm tên loại, trích yếu, Kính gửi"));
    if (!withTitle) return section;
    const radio = (value: "letter" | "named", label: string): HTMLElement => {
      const input = create("input", { type: "radio", name: "hf-variant", id: `hf-variant-${value}`, value, checked: variant === value });
      input.addEventListener("change", () => {
        variant = value;
        render();
      });
      return create("label", {}, input, ` ${label}`);
    };
    const text = (id: string, placeholder: string, value: string, onInput: (v: string) => void): HTMLInputElement => {
      const input = create("input", { type: "text", id, value, placeholder });
      input.addEventListener("input", () => onInput(input.value));
      return input;
    };
    section.append(radio("letter", "Công văn (V/v và Kính gửi)"), radio("named", "Văn bản có tên loại (quyết định, tờ trình...)"));
    if (variant === "named") {
      section.append(
        create("div", {}, text("hf-type-name", "Tên loại, ví dụ Quyết định", title.typeName, (v) => (title.typeName = v))),
        create("div", {}, text("hf-subject", "Trích yếu nội dung", title.subject, (v) => (title.subject = v))),
      );
    } else {
      const recipients = create("textarea", { id: "hf-recipients", rows: 3, placeholder: "Nơi nhận Kính gửi, mỗi nơi một dòng", value: title.recipients });
      recipients.addEventListener("input", () => {
        title.recipients = recipients.value;
      });
      section.append(
        create("div", {}, text("hf-subject", "Trích yếu (sau V/v)", title.subject, (v) => (title.subject = v))),
        create("div", {}, recipients),
      );
    }
    return section;
  }

  function render(): void {
    const profile = deps.currentStandard();
    body.replaceChildren();
    if (profile.headerKind === "dang") {
      body.append(
        create("p", { className: "note", textContent: "Chèn bảng cơ quan cấp trên, cơ quan ban hành (dấu *) và tiêu đề ĐẢNG CỘNG SẢN VIỆT NAM ở đầu tài liệu. Thể loại và trích yếu sẽ có ở bước TD3." }),
        parentSection(),
        issuerSection(profile),
        numberSection(),
      );
    } else {
      body.append(
        create("p", { className: "note", textContent: "Chèn bảng cơ quan, quốc hiệu và tiêu ngữ ở đầu tài liệu. Cỡ chữ, độ dài đường kẻ và lề lấy từ Tuỳ chọn nhanh." }),
        parentSection(),
        issuerSection(profile),
        numberSection(),
        titleSection(),
      );
    }
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
    collect,
  };
}
