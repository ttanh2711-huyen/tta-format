import { createLetter, type LetterContext } from "../office/createDocument";
import { isDocumentStandardSupported, readDocumentSnapshot } from "../office/documentStandard";
import { describeSnapshot, restoreSnapshot, takeSnapshot, type DocumentSnapshot } from "../office/snapshot";
import { summarizeReport } from "../pipeline/runner";
import type { FormatValues } from "../settings/formatOptions";
import type { StyleSpec } from "../ooxml/styleSpecs";
import { assembleDocumentBody } from "../ooxml/documentAssembler";
import { QUOC_HIEU } from "../ooxml/headerBlock";
import { findTemplate, visibleTemplates, type TemplateEntry } from "../templates/registry";
import { multiSignFromOptions, parseSignLines, validateMultiSign } from "../ooxml/multiSignBlock";
import { assembleMinutesBody, defaultMinutesParties } from "../templates/bienBan";
import { agreementNumberingXml, assembleAgreementBody, defaultAgreementParties, MAX_ARTICLES as MAX_AGREEMENT_ARTICLES, validateAgreement } from "../templates/hopDong";
import { partyRegulationFromOptions, assemblePartyRegulationBody, partyRegulationNumberingXml, validatePartyRegulation } from "../standards/hd05/templates/regulations";
import type { PartyNumbering, PartyTopLevel } from "../standards/hd05/numbering";
import { assemblePartyCongressBody, assemblePartyFormBody, assemblePartyJointBody, assemblePartyMinutesBody } from "../standards/hd05/templates/documents";
import { assembleSmallFormBody, findSmallForm, formSource } from "../templates/smallForms";
import { assembleRegulationBody, MAX_ARTICLES_PER_CHAPTER, MAX_CHAPTERS, regulationInputFromOptions, regulationNumberingXml, validateRegulation } from "../templates/quyChe";
import { assembleDecisionBody, decisionInputFromOptions, decisionKindInfo, decisionNumberingXml, MAX_ARTICLES, validateDecision } from "../templates/quyetDinh";
import { assemblePartyDocumentBody } from "../standards/hd05/assembleDocument";
import type { StandardId, StandardProfile } from "../standards/types";
import type { FooterCollect } from "./FooterForm";
import type { HeaderCollect, HeaderPreset } from "./HeaderForm";

export interface NewDocWizardDeps {
  currentStandard: () => StandardProfile;
  values: () => FormatValues;
  styleSpecs: () => StyleSpec[];
  behavior: () => { localFormatting: "keep" | "clear" };
  collectHeader: () => HeaderCollect;
  /** Đặt sẵn Khối đầu theo loại văn bản được chọn. */
  presetHeader?: (preset: HeaderPreset) => void;
  collectFooter: () => FooterCollect;
  /** Ghi chuẩn vào thuộc tính tài liệu sau khi tạo xong; bỏ trống nếu Word không hỗ trợ. */
  writeStandard?: (id: StandardId) => Promise<void>;
  /** Gọi sau khi tạo văn bản thành công (ví dụ ghi số văn bản vào sổ số). */
  afterCreate?: () => void;
}

type DocumentParts = Pick<LetterContext, "buildBody" | "headText" | "signerName" | "numberingXml" | "markerStyle">;

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

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/**
 * Nút "Tạo công văn" (T12): ghép các ô của Khối đầu và Khối cuối thành một công văn hoàn chỉnh trong tài liệu
 * hiện tại (thiết lập trang, cài style, dựng nội dung, đánh số trang). Tài liệu đã có nội dung thì hỏi trước,
 * và chụp bản gốc để khôi phục.
 */
export function mountNewDocWizard(container: HTMLElement, deps: NewDocWizardDeps): { refresh: () => void } {
  let busy = false;
  let snapshot: DocumentSnapshot | null = null;

  let entryKey = "cong-van";
  const decision = { authority: "", basis: "", articles: "3", khoan: "0", diem: "0" };
  const regulation = { chapters: "3", articles: "2", sections: false, attached: "", khoan: "1", diem: "0" };
  const signed = { lines: "", recipients: "", articles: "5" };
  const congress = { name: "" };
  const partyReg: { top: PartyTopLevel; numbering: PartyNumbering } = { top: "Chương", numbering: "roman" };
  const note = create("p", { className: "note" });
  const formBox = create("div", { id: "nw-form" });
  const create_ = button("Tạo công văn", () => void onCreate(false), { id: "nw-create" });
  const restore = button("Khôi phục bản trước khi tạo", () => onRestore(), { id: "nw-restore", hidden: true });
  const panel = create("div", { id: "nw-panel" });
  const message = create("p", { id: "nw-message" });
  message.setAttribute("role", "status");
  container.append(create("h2", { textContent: "Tạo công văn mới" }), note, formBox, create("div", {}, create_, " ", restore), panel, message);

  function show(kind: "ok" | "warning" | "error" | "", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  function setBusy(value: boolean): void {
    busy = value;
    create_.disabled = value;
    restore.disabled = value;
  }

  function refresh(): void {
    const profile = deps.currentStandard();
    const party = profile.headerKind === "dang";
    note.textContent = `Điền Khối đầu và Khối cuối bên dưới rồi bấm: add-in thiết lập trang${party ? " (lề đối xứng nếu bật in hai mặt)" : ""}, cài style, dựng khối đầu, ${party ? "thể loại hoặc Kính gửi, căn cứ" : "Kính gửi"}, chỗ nội dung, khối cuối và đánh số trang. Sau đó bạn gõ nội dung vào chỗ con trỏ.`;
    if (!visibleTemplates(profile).some((item) => item.key === entryKey && item.supported)) entryKey = "cong-van";
    const entry = findTemplate(profile, entryKey);
    create_.textContent = party ? "Tạo văn bản Đảng" : entry?.family === "congVan" || !entry ? "Tạo công văn" : `Tạo ${entry.label.toLowerCase()}`;
    renderForm(profile);
    panel.replaceChildren();
    show("", "");
  }

  /** Chọn loại văn bản từ registry (T17) và các ô riêng của nghị quyết, quyết định cá biệt. */
  function renderForm(profile: StandardProfile): void {
    formBox.replaceChildren();
    const entries = visibleTemplates(profile);
    const select = create("select", { id: "nw-form-kind" });
    for (const entry of entries) {
      const option = create("option", { value: entry.key, textContent: entry.supported ? entry.label : `${entry.label} — chưa hỗ trợ: ${entry.reason ?? ""}` });
      option.disabled = !entry.supported;
      select.append(option);
    }
    select.value = entryKey;
    select.addEventListener("change", () => {
      const chosen = findTemplate(profile, select.value);
      if (!chosen?.supported) return;
      entryKey = chosen.key;
      applyPreset(chosen);
      refresh();
    });
    formBox.append(create("label", {}, "Loại văn bản ", select));
    const entry = findTemplate(profile, entryKey);
    const field =
      (el: HTMLInputElement | HTMLTextAreaElement, onInput: (v: string) => void): HTMLElement => {
        el.addEventListener("input", () => onInput(el.value));
        return create("div", {}, el);
      };
    if (entry?.family === "form") {
      const smallForm = findSmallForm(entry.docTypeId);
      formBox.append(create("p", { className: "note", textContent: `Khối đầu đã chuyển sang thể loại và trích yếu; nhập trích yếu. Add-in dựng các mục cần điền (${smallForm?.fields.join(" ")}) rồi khối cuối. Lưu ý: ${smallForm ? formSource(smallForm, profile.headerKind) : ""}.` }));
      return;
    }
    if (entry?.family === "congress") {
      formBox.append(
        create("p", { className: "note", textContent: 'Khối đầu đã chuyển sang thể loại và trích yếu; nhập thể loại (ví dụ Nghị quyết) và trích yếu. Add-in thêm dòng tên đại hội in hoa đậm trên thể loại và gợi ý ký hiệu "ĐH". Quyền hạn ký thường là T/M Đoàn chủ tịch; dòng xác nhận chữ ký nhập ở Khối cuối.' }),
        field(create("input", { type: "text", id: "nw-congress", value: congress.name, placeholder: "Tên đại hội, ví dụ Đại hội đại biểu Đảng bộ tỉnh lần thứ XI" }), (v) => (congress.name = v)),
      );
      return;
    }
    if (entry?.family === "minutes" || entry?.family === "agreement" || entry?.family === "joint") {
      const party = profile.headerKind === "dang";
      const defaults = entry.family === "agreement" ? defaultAgreementParties(2) : entry.family === "joint" ? [{ title: "T/M cơ quan thứ nhất" }, { title: "T/M cơ quan thứ hai" }] : party ? [{ title: "Người ghi biên bản" }, { title: "Chủ trì hội nghị" }] : defaultMinutesParties();
      if (signed.lines.trim() === "") signed.lines = defaults.map((p) => `${p.title} | |`).join("\n");
      const lines = create("textarea", { id: "nw-sign-lines", rows: 4, value: signed.lines, placeholder: party ? "Mỗi bên một dòng: Tên bên | Chức vụ | Họ tên | Xác nhận chữ ký (nếu không đóng dấu)" : "Mỗi bên một dòng: Tên bên | Chức vụ | Họ tên (2 đến 4 bên)" });
      const recipients = create("textarea", { id: "nw-sign-recipients", rows: 3, value: signed.recipients, placeholder: "Nơi nhận, mỗi nơi một dòng (bỏ trống nếu không cần)" });
      formBox.append(
        create("p", { className: "note", textContent: entry.family === "joint" ? "Khối đầu: nhập mỗi cơ quan ban hành một dòng và nối chữ viết tắt cơ quan bằng dấu gạch nối. Add-in dựng thể loại, trích yếu, chỗ nội dung và đúng hai khối ký." : entry.family === "minutes" ? 'Khối đầu đã chuyển sang thể loại hoặc tên loại; nhập trích yếu (ví dụ "Về việc ..."). Add-in dựng các mục Thời gian, Địa điểm, Thành phần, Chủ trì, Thư ký, Nội dung và khối ký nhiều bên.' : 'Khối đầu đã chuyển sang "Văn bản có tên loại"; nhập trích yếu. Phụ lục III không có mẫu cho loại này: add-in dùng bố cục đề xuất (các bên, các Điều, khối ký nhiều bên, không có "Nơi nhận" cố định); hãy đối chiếu trước khi dùng.' }),
        field(lines, (v) => (signed.lines = v)),
        field(recipients, (v) => (signed.recipients = v)),
      );
      if (entry.family === "agreement") {
        formBox.append(field(create("input", { type: "text", id: "nw-agreement-articles", value: signed.articles, placeholder: `Số Điều (1 đến ${MAX_AGREEMENT_ARTICLES})`, inputMode: "numeric" }), (v) => (signed.articles = v)));
      }
      return;
    }
    if (entry?.family === "regulation" && profile.headerKind === "dang") {
      const sections = create("input", { type: "checkbox", id: "nw-sections", checked: regulation.sections });
      sections.addEventListener("change", () => {
        regulation.sections = sections.checked;
      });
      const top = create("select", { id: "nw-party-top" });
      for (const value of ["Chương", "Phần"] as const) top.append(create("option", { value, textContent: value }));
      top.value = partyReg.top;
      top.addEventListener("change", () => {
        partyReg.top = top.value as PartyTopLevel;
      });
      const numbering = create("select", { id: "nw-party-numbering" });
      numbering.append(create("option", { value: "roman", textContent: "Số La Mã (Chương I)" }), create("option", { value: "words", textContent: "Số bằng chữ (Chương thứ hai)" }));
      numbering.value = partyReg.numbering;
      numbering.addEventListener("change", () => {
        partyReg.numbering = numbering.value as PartyNumbering;
      });
      formBox.append(
        create("p", { className: "note", textContent: `Khối đầu đã chuyển sang thể loại và trích yếu. Add-in dựng các Phần hoặc Chương (nhãn tự đánh số, tên in hoa đậm ở dòng dưới), Mục, Điều (số chạy liên tục) và Khoản; con trỏ ở tên Phần hoặc Chương thứ nhất. Căn cứ nhập ở Khối đầu.` }),
        create("div", {}, create("label", {}, "Cấp cao nhất ", top), " ", create("label", {}, "Cách ghi số ", numbering)),
        field(create("input", { type: "text", id: "nw-chapters", value: regulation.chapters, placeholder: `Số Phần hoặc Chương (1 đến ${MAX_CHAPTERS})`, inputMode: "numeric" }), (v) => (regulation.chapters = v)),
        field(create("input", { type: "text", id: "nw-reg-articles", value: regulation.articles, placeholder: `Số Điều mỗi Phần hoặc Chương (1 đến ${MAX_ARTICLES_PER_CHAPTER})`, inputMode: "numeric" }), (v) => (regulation.articles = v)),
        field(create("input", { type: "text", id: "nw-reg-khoan", value: regulation.khoan, placeholder: "Số Khoản dựng sẵn mỗi Điều (0 đến 10)", inputMode: "numeric" }), (v) => (regulation.khoan = v)),
        field(create("input", { type: "text", id: "nw-reg-diem", value: regulation.diem, placeholder: "Số Điểm dựng sẵn mỗi Khoản (0 đến 10)", inputMode: "numeric" }), (v) => (regulation.diem = v)),
        create("label", {}, sections, " Mỗi Phần hoặc Chương có một Mục"),
      );
      return;
    }
    if (entry?.family === "regulation") {
      const sections = create("input", { type: "checkbox", id: "nw-sections", checked: regulation.sections });
      sections.addEventListener("change", () => {
        regulation.sections = sections.checked;
      });
      formBox.append(
        create("p", { className: "note", textContent: `Khối đầu đã chuyển sang "Văn bản có tên loại"; nhập trích yếu (ví dụ "Làm việc của ..."). Add-in dựng các Chương (nhãn tự đánh số La Mã, tên Chương in hoa đậm), các Điều (số chạy liên tục) và Khoản; con trỏ ở tên Chương thứ nhất.` }),
        field(create("input", { type: "text", id: "nw-attached", value: regulation.attached, placeholder: "Dòng ban hành kèm theo, ví dụ Ban hành kèm theo Quyết định số ... ngày ... của ..." }), (v) => (regulation.attached = v)),
        field(create("input", { type: "text", id: "nw-chapters", value: regulation.chapters, placeholder: `Số Chương (1 đến ${MAX_CHAPTERS})`, inputMode: "numeric" }), (v) => (regulation.chapters = v)),
        field(create("input", { type: "text", id: "nw-reg-articles", value: regulation.articles, placeholder: `Số Điều mỗi Chương (1 đến ${MAX_ARTICLES_PER_CHAPTER})`, inputMode: "numeric" }), (v) => (regulation.articles = v)),
        field(create("input", { type: "text", id: "nw-reg-khoan", value: regulation.khoan, placeholder: "Số Khoản dựng sẵn mỗi Điều (0 đến 10)", inputMode: "numeric" }), (v) => (regulation.khoan = v)),
        field(create("input", { type: "text", id: "nw-reg-diem", value: regulation.diem, placeholder: "Số Điểm dựng sẵn mỗi Khoản (0 đến 10)", inputMode: "numeric" }), (v) => (regulation.diem = v)),
        create("label", {}, sections, " Mỗi Chương có một Mục"),
      );
      return;
    }
    if (entry?.family !== "decision" || !entry.decisionKind) return;
    const info = decisionKindInfo(entry.decisionKind);
    formBox.append(
      create("p", { className: "note", textContent: `Khối đầu đã chuyển sang "Văn bản có tên loại"; nhập trích yếu (ví dụ "Về việc ..."). Add-in tự ghi tên loại "${info.typeName.toUpperCase()}". Phần dưới tạo thẩm quyền ban hành, căn cứ, dòng "${info.verdict}" và các Điều đánh số tự động.` }),
      field(create("input", { type: "text", id: "nw-authority", value: decision.authority, placeholder: "Thẩm quyền ban hành, ví dụ Chủ tịch Ủy ban nhân dân thành phố" }), (v) => (decision.authority = v)),
      field(create("textarea", { id: "nw-basis", rows: 4, value: decision.basis, placeholder: "Căn cứ, mỗi căn cứ một dòng" }), (v) => (decision.basis = v)),
      field(create("input", { type: "text", id: "nw-articles", value: decision.articles, placeholder: `Số Điều (1 đến ${MAX_ARTICLES})`, inputMode: "numeric" }), (v) => (decision.articles = v)),
      field(create("input", { type: "text", id: "nw-khoan", value: decision.khoan, placeholder: "Số Khoản dựng sẵn mỗi Điều (0 đến 10)", inputMode: "numeric" }), (v) => (decision.khoan = v)),
      field(create("input", { type: "text", id: "nw-diem", value: decision.diem, placeholder: "Số Điểm dựng sẵn mỗi Khoản (0 đến 10)", inputMode: "numeric" }), (v) => (decision.diem = v)),
    );
  }

  /** Đưa loại vừa chọn vào Khối đầu: dạng khối, tên loại, chữ viết tắt gợi ý. */
  function applyPreset(entry: TemplateEntry): void {
    const form = entry.family === "congVan" ? "congVan" : entry.family === "toTrinh" ? "toTrinh" : "titled";
    const typeName = entry.family === "decision" && entry.decisionKind ? decisionKindInfo(entry.decisionKind).typeName : entry.typeName;
    deps.presetHeader?.({ form, typeName, abbr: entry.abbr });
  }

  async function hasContent(): Promise<boolean> {
    return isDocumentStandardSupported() ? (await readDocumentSnapshot()).hasContent : false;
  }

  async function onCreate(confirmed: boolean): Promise<void> {
    if (busy) return;
    const header = deps.collectHeader();
    const footer = deps.collectFooter();
    const profile = deps.currentStandard();
    const entry = findTemplate(profile, entryKey);
    // Biên bản và văn bản nhiều bên ký có khối ký riêng nên không dùng Khối cuối.
    const multiSign = entry?.family === "minutes" || entry?.family === "agreement" || entry?.family === "joint";
    if (!header.ok || (!multiSign && !footer.ok)) {
      show("error", [header.ok ? "" : `Khối đầu: ${header.error}`, footer.ok ? "" : `Khối cuối: ${footer.error}`].filter(Boolean).join(" "));
      return;
    }
    let parts: DocumentParts;
    if (multiSign && header.ok && header.kind === "quocHieu") {
      if (header.titleInput?.variant !== "named") {
        show("error", "Ở Khối đầu hãy chọn \"Văn bản có tên loại\" và nhập trích yếu.");
        return;
      }
      const parties = parseSignLines(signed.lines);
      const sign = multiSignFromOptions(profile, deps.values(), { parties, recipients: signed.recipients.split(/\r?\n/), luu: "" });
      const document = { header: header.input, title: { ...header.titleInput, typeName: entry.typeName } };
      const signProblems = validateMultiSign(sign);
      if (signProblems.length > 0) {
        show("error", signProblems.join(" "));
        return;
      }
      if (entry.family === "minutes") {
        parts = { buildBody: () => assembleMinutesBody(document, profile.styles.prefix, sign), headText: QUOC_HIEU, signerName: "" };
      } else {
        const agreement = { parties: parties.length, articleCount: Number(signed.articles), stylePrefix: profile.styles.prefix, firstLineIndentMm: typeof deps.values().firstLineIndent === "number" ? (deps.values().firstLineIndent as number) : profile.styles.defaults.firstLineIndentMm };
        const problems = [...validateAgreement(agreement)];
        if (problems.length > 0) {
          show("error", problems.join(" "));
          return;
        }
        parts = { buildBody: () => assembleAgreementBody(document, agreement, sign), headText: QUOC_HIEU, signerName: "", numberingXml: agreementNumberingXml(agreement) };
      }
    } else if (multiSign && header.ok && header.kind === "dang") {
      if (!header.titleInput || header.titleInput.variant === "congVan") {
        show("error", "Ở Khối đầu hãy chọn thể loại (không phải công văn) và nhập trích yếu.");
        return;
      }
      const sign = multiSignFromOptions(profile, deps.values(), { parties: parseSignLines(signed.lines), recipients: signed.recipients.split(/\r?\n/), luu: "" });
      const signProblems = validateMultiSign(sign);
      if (signProblems.length > 0) {
        show("error", signProblems.join(" "));
        return;
      }
      const document = { header: header.input, title: { ...header.titleInput, typeName: entry.family === "joint" ? header.titleInput.typeName : entry.typeName } };
      if (entry.family === "joint") parts = { buildBody: (body) => assemblePartyJointBody(document, sign, body), headText: profile.identifyingTitle, signerName: "" };
      else parts = { buildBody: () => assemblePartyMinutesBody(document, profile.styles.prefix, sign), headText: profile.identifyingTitle, signerName: "" };
    } else if (!footer.ok) {
      return;
    } else if (header.kind === "dang" && footer.kind === "dang" && entry?.family === "congress") {
      if (!header.titleInput || header.titleInput.variant === "congVan") {
        show("error", "Ở Khối đầu hãy chọn thể loại (không phải công văn) và nhập trích yếu.");
        return;
      }
      if (congress.name.trim() === "") {
        show("error", "Hãy nhập tên đại hội.");
        return;
      }
      const document = { header: header.input, title: header.titleInput, footer: footer.input };
      const size = typeof deps.values().fontSize === "number" ? (deps.values().fontSize as number) : profile.styles.defaults.fontSizePt;
      parts = { buildBody: (body) => assemblePartyCongressBody(document, congress.name, size, body), headText: profile.identifyingTitle, signerName: footer.input.signerName };
    } else if (header.kind === "dang" && footer.kind === "dang" && entry?.family === "form") {
      const smallForm = findSmallForm(entry.docTypeId);
      if (!smallForm || !header.titleInput || header.titleInput.variant === "congVan") {
        show("error", "Ở Khối đầu hãy chọn thể loại (không phải công văn) và nhập trích yếu.");
        return;
      }
      const document = { header: header.input, title: { ...header.titleInput, typeName: entry.typeName }, footer: footer.input };
      parts = { buildBody: () => assemblePartyFormBody(document, smallForm, profile.styles.prefix), headText: profile.identifyingTitle, signerName: footer.input.signerName };
    } else if (header.kind === "quocHieu" && footer.kind === "quocHieu" && entry?.family === "form") {
      const smallForm = findSmallForm(entry.docTypeId);
      if (!smallForm || header.titleInput?.variant !== "named") {
        show("error", "Ở Khối đầu hãy chọn \"Văn bản có tên loại\" và nhập trích yếu.");
        return;
      }
      const document = { header: header.input, title: { ...header.titleInput, typeName: entry.typeName }, footer: footer.input };
      parts = { buildBody: () => assembleSmallFormBody(document, smallForm, profile.styles.prefix), headText: QUOC_HIEU, signerName: footer.input.signerName };
    } else if (header.kind === "quocHieu" && footer.kind === "quocHieu" && entry?.family === "regulation") {
      if (header.titleInput?.variant !== "named") {
        show("error", "Ở Khối đầu hãy chọn \"Văn bản có tên loại\" và nhập trích yếu.");
        return;
      }
      const regulationInput = regulationInputFromOptions(profile, deps.values(), { chapters: Number(regulation.chapters), articlesPerChapter: Number(regulation.articles), withSections: regulation.sections, attachedLine: regulation.attached, khoanPerDieu: Number(regulation.khoan), diemPerKhoan: Number(regulation.diem) });
      const regulationProblems = validateRegulation(regulationInput);
      if (regulationProblems.length > 0) {
        show("error", regulationProblems.join(" "));
        return;
      }
      const document = { header: header.input, title: { ...header.titleInput, typeName: entry.typeName }, footer: footer.input };
      parts = {
        buildBody: () => assembleRegulationBody(document, regulationInput),
        headText: QUOC_HIEU,
        signerName: footer.input.signerName,
        numberingXml: regulationNumberingXml(regulationInput),
        markerStyle: `${profile.styles.prefix}TieuDePhanChuong`,
      };
    } else if (header.kind === "quocHieu" && footer.kind === "quocHieu" && entry?.family === "decision" && entry.decisionKind) {
      if (header.titleInput?.variant !== "named") {
        show("error", "Ở Khối đầu hãy chọn \"Văn bản có tên loại\" và nhập trích yếu.");
        return;
      }
      const info = decisionKindInfo(entry.decisionKind);
      const decisionInput = decisionInputFromOptions(profile, deps.values(), { kind: entry.decisionKind, authority: decision.authority, basis: decision.basis.split(/\r?\n/), articleCount: Number(decision.articles), khoanPerDieu: Number(decision.khoan), diemPerKhoan: Number(decision.diem) });
      const decisionProblems = validateDecision(decisionInput);
      if (decisionProblems.length > 0) {
        show("error", decisionProblems.join(" "));
        return;
      }
      const document = { header: header.input, title: { ...header.titleInput, typeName: info.typeName }, footer: footer.input };
      parts = { buildBody: () => assembleDecisionBody(document, decisionInput), headText: QUOC_HIEU, signerName: footer.input.signerName, numberingXml: decisionNumberingXml(decisionInput) };
    } else if (header.kind === "quocHieu" && footer.kind === "quocHieu") {
      const document = { header: header.input, title: header.titleInput, footer: footer.input };
      parts = { buildBody: (body) => assembleDocumentBody(document, body), headText: QUOC_HIEU, signerName: footer.input.signerName };
    } else if (header.kind === "dang" && footer.kind === "dang" && entry?.family === "regulation") {
      if (!header.titleInput || header.titleInput.variant === "congVan") {
        show("error", "Ở Khối đầu hãy chọn thể loại (không phải công văn) và nhập trích yếu.");
        return;
      }
      const reg = partyRegulationFromOptions(profile, deps.values(), { chapters: Number(regulation.chapters), articlesPerChapter: Number(regulation.articles), withSections: regulation.sections, top: partyReg.top, numbering: partyReg.numbering, khoanPerDieu: Number(regulation.khoan), diemPerKhoan: Number(regulation.diem) });
      const regProblems = validatePartyRegulation(reg);
      if (regProblems.length > 0) {
        show("error", regProblems.join(" "));
        return;
      }
      const document = { header: header.input, title: { ...header.titleInput, typeName: entry.typeName }, footer: footer.input };
      parts = {
        buildBody: () => assemblePartyRegulationBody(document, reg),
        headText: profile.identifyingTitle,
        signerName: footer.input.signerName,
        numberingXml: partyRegulationNumberingXml(reg),
        markerStyle: `${profile.styles.prefix}TenPhanChuong`,
      };
    } else if (header.kind === "dang" && footer.kind === "dang") {
      const document = { header: header.input, title: header.titleInput, footer: footer.input };
      parts = { buildBody: (body) => assemblePartyDocumentBody(document, body), headText: profile.identifyingTitle, signerName: footer.input.signerName };
    } else {
      show("error", "Khối đầu và Khối cuối không cùng một chuẩn; hãy tải lại add-in.");
      return;
    }
    setBusy(true);
    try {
      if (!confirmed) {
        const reasons: string[] = [];
        if (await hasContent()) reasons.push("Tài liệu đang có nội dung; công văn mới sẽ thay toàn bộ nội dung đó (có thể khôi phục).");
        if (footer.ok && footer.academicTitle) reasons.push(`Họ tên người ký có "${footer.academicTitle}", quy định chỉ ghi họ và tên.`);
        if (reasons.length > 0) {
          const box = create("div", { className: "nw-box", id: "nw-confirm" });
          box.setAttribute("role", "alert");
          box.append(
            ...reasons.map((reason) => create("p", { textContent: reason })),
            button("Vẫn tạo", () => void onCreate(true)),
            button("Huỷ", () => panel.replaceChildren()),
          );
          panel.replaceChildren(box);
          show("", "");
          return;
        }
      }
      panel.replaceChildren();
      const context: LetterContext = {
        profile,
        values: deps.values(),
        specs: deps.styleSpecs(),
        localFormatting: deps.behavior().localFormatting,
        ...parts,
      };
      show("ok", "Đang chụp bản trước khi tạo…");
      let taken: DocumentSnapshot;
      try {
        taken = await takeSnapshot();
      } catch (error) {
        show("error", `Không chụp được bản trước khi tạo nên chưa làm gì, tài liệu giữ nguyên: ${errorText(error)}`);
        return;
      }
      const report = await createLetter(context, (event) => {
        if (event.phase === "start") show("ok", `Đang chạy bước ${event.stepId}…`);
      });
      if (report.changed) {
        snapshot = taken;
        restore.hidden = false;
      }
      let standardNote = "";
      if (report.ok) deps.afterCreate?.();
      if (report.ok && deps.writeStandard) {
        try {
          await deps.writeStandard(profile.id);
        } catch (error) {
          standardNote = ` Chưa ghi được chuẩn vào tài liệu: ${errorText(error)}`;
        }
      }
      const summary = summarizeReport(report);
      const box = create("div", { className: "nw-box", id: "nw-report" }, create("strong", { textContent: report.ok ? "Kết quả" : "Chưa hoàn tất" }));
      const list = create("ul");
      for (const line of summary.lines) list.append(create("li", { textContent: line }));
      box.append(list);
      panel.replaceChildren(box);
      show(summary.kind, report.ok ? `Đã tạo văn bản; hãy gõ nội dung vào chỗ con trỏ.${standardNote}` : "Có bước bị lỗi, xem chi tiết bên trên; có thể khôi phục bản trước khi tạo.");
    } catch (error) {
      show("error", `Lỗi ngoài dự kiến: ${errorText(error)}`);
    } finally {
      setBusy(false);
    }
  }

  function onRestore(): void {
    if (busy || !snapshot) return;
    const saved = snapshot;
    const box = create("div", { className: "nw-box", id: "nw-restore-confirm" });
    box.setAttribute("role", "alert");
    box.append(
      create("p", { textContent: `Khôi phục sẽ thay lại thân văn bản, thiết lập trang và header bằng ${describeSnapshot(saved)}. Mọi chỉnh sửa sau thời điểm đó sẽ mất.` }),
      button("Khôi phục", () => {
        void (async () => {
          setBusy(true);
          panel.replaceChildren();
          try {
            const result = await restoreSnapshot(saved);
            if (result.problems.length === 0) {
              snapshot = null;
              restore.hidden = true;
              show("ok", `Đã khôi phục về ${describeSnapshot(saved)}.`);
            } else {
              show("warning", `Đã khôi phục nhưng còn điểm chưa khớp: ${result.problems.join("; ")}.`);
            }
          } catch (error) {
            show("error", `Không khôi phục được: ${errorText(error)}`);
          } finally {
            setBusy(false);
          }
        })();
      }),
      button("Huỷ", () => panel.replaceChildren()),
    );
    panel.replaceChildren(box);
  }

  refresh();
  return { refresh };
}
