import { buildRebuiltHeader, normalizeDate, normalizeSo, validateHeaderFields, type HeaderFields } from "../checker/headerRebuild";
import { previewHeaderRebuild, rebuildHeader } from "../office/rebuildHeader";
import { describeSnapshot, restoreSnapshot, takeSnapshot, type DocumentSnapshot } from "../office/snapshot";
import type { FormatValues } from "../settings/formatOptions";
import { STANDARDS } from "../standards";
import type { StandardProfile } from "../standards/types";

function create<K extends keyof HTMLElementTagNameMap>(tag: K, props: Partial<HTMLElementTagNameMap[K]> = {}, ...children: Array<Node | string>): HTMLElementTagNameMap[K] {
  const element = document.createElement(tag);
  Object.assign(element, props);
  element.append(...children);
  return element;
}

const errorText = (error: unknown): string => (error instanceof Error ? error.message : String(error));
const lines = (value: string): string[] => value.split(/\r?\n/).map((line) => line.trim()).filter((line) => line !== "");

/**
 * Mục "Sửa khối đầu": đọc khối đầu đang có (kể cả gõ tay, sai quy định), tách nội dung ra các ô để xem và sửa, rồi xoá
 * khối cũ và dựng khối đầu mới đúng quy định của chuẩn đang chọn (bảng khối đầu, tên loại, trích yếu, Kính gửi).
 */
export function mountHeaderRebuild(container: HTMLElement, deps: { currentStandard: () => StandardProfile; values: () => FormatValues }): { refresh: () => void } {
  let busy = false;
  let snapshot: DocumentSnapshot | null = null;
  const read = create("button", { type: "button", id: "hr-read", textContent: "Đọc khối đầu hiện có", className: "btn-primary" });
  const form = create("div", { id: "hr-form" });
  const message = create("p", { id: "hr-message" });
  message.setAttribute("role", "status");
  container.append(
    create("p", { className: "note", textContent: "Dùng cho văn bản đã có khối đầu (tên cơ quan, quốc hiệu, số, ngày tháng, tên loại, trích yếu, Kính gửi) nhưng trình bày chưa đúng quy định. Add-in lấy nội dung cũ, chuẩn hoá (số, ngày tháng, in hoa, đường kẻ, cỡ chữ) rồi dựng lại khối đầu mới thay cho khối cũ; phần thân giữ nguyên." }),
    create("div", { className: "rp-actions" }, read),
    form,
    message,
  );

  function show(kind: "ok" | "warning" | "error" | "", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  function field(id: string, label: string, control: HTMLInputElement | HTMLTextAreaElement): HTMLElement {
    control.id = id;
    return create("label", { className: "fld" }, create("span", { textContent: label }), control);
  }

  function renderForm(fields: HeaderFields, notes: readonly string[]): void {
    const party = deps.currentStandard().headerKind === "dang";
    const parent = create("textarea", { rows: 2, value: fields.parentLines.join("\n") });
    const issuer = create("textarea", { rows: 2, value: fields.issuerLines.join("\n") });
    const so = create("input", { type: "text", value: fields.so });
    const placeDate = create("input", { type: "text", value: fields.placeDate });
    const typeName = create("input", { type: "text", value: fields.typeName, placeholder: "Báo cáo, Tờ trình, Hướng dẫn…" });
    typeName.setAttribute("list", "hr-types");
    const types = create("datalist", { id: "hr-types" }, ...[...new Set(STANDARDS.flatMap((standard) => standard.docTypes.map((type) => type.name)))].map((name) => create("option", { value: name })));
    const subject = create("textarea", { rows: 2, value: fields.subject });
    const recipients = create("textarea", { rows: 2, value: fields.recipients.join("\n"), placeholder: "Mỗi nơi một dòng" });
    const letter = create("input", { type: "radio", name: "hr-kind", id: "hr-kind-letter", checked: fields.kind === "letter" });
    const named = create("input", { type: "radio", name: "hr-kind", id: "hr-kind-named", checked: fields.kind === "named" });
    const typeRow = field("hr-type", "Tên loại văn bản", typeName);
    const syncKind = (): void => {
      typeRow.hidden = letter.checked;
    };
    letter.addEventListener("change", syncKind);
    named.addEventListener("change", syncKind);
    syncKind();

    const apply = create("button", { type: "button", id: "hr-apply", textContent: "Dựng lại khối đầu đúng quy định", className: "btn-primary" });
    const restore = create("button", { type: "button", id: "hr-restore", textContent: "Khôi phục bản gốc", hidden: snapshot === null });
    apply.addEventListener("click", () =>
      void guarded(async () => {
        const edited: HeaderFields = {
          parentLines: lines(parent.value),
          issuerLines: lines(issuer.value),
          so: so.value.trim() === "" ? "" : normalizeSo(so.value, deps.currentStandard()),
          placeDate: placeDate.value.trim() === "" ? "" : normalizeDate(placeDate.value),
          kind: letter.checked ? "letter" : "named",
          typeName: typeName.value.trim(),
          subject: subject.value.trim(),
          recipients: lines(recipients.value),
        };
        const problems = validateHeaderFields(edited);
        if (problems.length > 0) {
          show("error", problems.join(" "));
          return;
        }
        const profile = deps.currentStandard();
        const xml = buildRebuiltHeader(edited, profile, deps.values());
        show("ok", "Đang chụp bản gốc để có thể khôi phục…");
        const taken = await takeSnapshot();
        const result = await rebuildHeader(profile, xml);
        snapshot = taken;
        restore.hidden = false;
        show(result.verified ? "ok" : "warning", result.verified ? `Đã dựng lại khối đầu đúng quy định ${profile.name}. Muốn quay lại thì bấm Khôi phục bản gốc.` : "Đã dựng lại nhưng đọc lại không thấy tiêu đề trong bảng đầu tiên; hãy kiểm tra văn bản.");
      }),
    );
    restore.addEventListener("click", () =>
      void guarded(async () => {
        if (!snapshot) return;
        const result = await restoreSnapshot(snapshot);
        show(result.problems.length > 0 ? "warning" : "ok", result.problems.length > 0 ? `Đã khôi phục nhưng: ${result.problems.join("; ")}.` : `Đã khôi phục về ${describeSnapshot(snapshot)}.`);
        snapshot = null;
        restore.hidden = true;
      }),
    );

    form.replaceChildren(
      ...notes.map((note) => create("p", { className: "warning", textContent: note })),
      field("hr-parent", party ? "Cơ quan cấp trên (mỗi dòng một dòng in)" : "Cơ quan chủ quản (bỏ trống nếu không có)", parent),
      field("hr-issuer", "Cơ quan ban hành", issuer),
      field("hr-so", "Số, ký hiệu", so),
      field("hr-date", "Địa danh, ngày tháng", placeDate),
      create("div", { className: "fld" }, create("label", {}, letter, " Công văn"), " ", create("label", {}, named, " Có tên loại")),
      typeRow,
      types,
      field("hr-subject", "Trích yếu nội dung", subject),
      field("hr-recipients", "Kính gửi (nếu có)", recipients),
      create("div", { className: "rp-actions" }, apply, restore),
    );
  }

  async function guarded(task: () => Promise<void>): Promise<void> {
    if (busy) return;
    busy = true;
    read.disabled = true;
    try {
      await task();
    } catch (error) {
      show("error", `Không làm được: ${errorText(error)}`);
    } finally {
      busy = false;
      read.disabled = false;
    }
  }

  read.addEventListener("click", () =>
    void guarded(async () => {
      show("ok", "Đang đọc khối đầu…");
      const plan = await previewHeaderRebuild(deps.currentStandard());
      if ("error" in plan) {
        form.replaceChildren();
        show("error", plan.error);
        return;
      }
      renderForm(plan.fields, plan.notes);
      show("ok", `Đã đọc khối đầu (${plan.end + 1} đoạn từ đầu văn bản sẽ được thay). Xem lại các ô, sửa nếu cần, rồi bấm "Dựng lại khối đầu đúng quy định".`);
    }),
  );

  return {
    refresh: () => {
      form.replaceChildren();
      show("", "");
    },
  };
}
