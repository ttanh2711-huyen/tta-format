import { createLetter, type LetterContext } from "../office/createDocument";
import { isDocumentStandardSupported, readDocumentSnapshot } from "../office/documentStandard";
import { describeSnapshot, restoreSnapshot, takeSnapshot, type DocumentSnapshot } from "../office/snapshot";
import { summarizeReport } from "../pipeline/runner";
import type { FormatValues } from "../settings/formatOptions";
import type { StyleSpec } from "../ooxml/styleSpecs";
import { assembleDocumentBody } from "../ooxml/documentAssembler";
import { QUOC_HIEU } from "../ooxml/headerBlock";
import { findTemplate, templatesFor, type TemplateEntry } from "../templates/registry";
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
}

type DocumentParts = Pick<LetterContext, "buildBody" | "headText" | "signerName" | "numberingXml">;

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
  const decision = { authority: "", basis: "", articles: "3" };
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
    if (!findTemplate(profile, entryKey)?.supported) entryKey = "cong-van";
    const entry = findTemplate(profile, entryKey);
    create_.textContent = party ? "Tạo văn bản Đảng" : entry?.family === "congVan" || !entry ? "Tạo công văn" : `Tạo ${entry.label.toLowerCase()}`;
    renderForm(profile);
    panel.replaceChildren();
    show("", "");
  }

  /** Chọn loại văn bản từ registry (T17) và các ô riêng của nghị quyết, quyết định cá biệt. */
  function renderForm(profile: StandardProfile): void {
    formBox.replaceChildren();
    const entries = templatesFor(profile);
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
    if (entry?.family !== "decision" || !entry.decisionKind) return;
    const info = decisionKindInfo(entry.decisionKind);
    const field = (el: HTMLInputElement | HTMLTextAreaElement, onInput: (v: string) => void): HTMLElement => {
      el.addEventListener("input", () => onInput(el.value));
      return create("div", {}, el);
    };
    formBox.append(
      create("p", { className: "note", textContent: `Khối đầu đã chuyển sang "Văn bản có tên loại"; nhập trích yếu (ví dụ "Về việc ..."). Add-in tự ghi tên loại "${info.typeName.toUpperCase()}". Phần dưới tạo thẩm quyền ban hành, căn cứ, dòng "${info.verdict}" và các Điều đánh số tự động.` }),
      field(create("input", { type: "text", id: "nw-authority", value: decision.authority, placeholder: "Thẩm quyền ban hành, ví dụ Chủ tịch Ủy ban nhân dân thành phố" }), (v) => (decision.authority = v)),
      field(create("textarea", { id: "nw-basis", rows: 4, value: decision.basis, placeholder: "Căn cứ, mỗi căn cứ một dòng" }), (v) => (decision.basis = v)),
      field(create("input", { type: "text", id: "nw-articles", value: decision.articles, placeholder: `Số Điều (1 đến ${MAX_ARTICLES})`, inputMode: "numeric" }), (v) => (decision.articles = v)),
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
    if (!header.ok || !footer.ok) {
      show("error", [header.ok ? "" : `Khối đầu: ${header.error}`, footer.ok ? "" : `Khối cuối: ${footer.error}`].filter(Boolean).join(" "));
      return;
    }
    const profile = deps.currentStandard();
    let parts: DocumentParts;
    const entry = findTemplate(profile, entryKey);
    if (header.kind === "quocHieu" && footer.kind === "quocHieu" && entry?.family === "decision" && entry.decisionKind) {
      if (header.titleInput?.variant !== "named") {
        show("error", "Ở Khối đầu hãy chọn \"Văn bản có tên loại\" và nhập trích yếu.");
        return;
      }
      const info = decisionKindInfo(entry.decisionKind);
      const decisionInput = decisionInputFromOptions(profile, deps.values(), { kind: entry.decisionKind, authority: decision.authority, basis: decision.basis.split(/\r?\n/), articleCount: Number(decision.articles) });
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
        if (footer.academicTitle) reasons.push(`Họ tên người ký có "${footer.academicTitle}", quy định chỉ ghi họ và tên.`);
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
