import type { TemplatePlan } from "../checker/templatePlan";
import { makeTemplate, previewTemplate } from "../office/makeTemplate";
import { describeSnapshot, restoreSnapshot, takeSnapshot, type DocumentSnapshot } from "../office/snapshot";
import type { StandardProfile } from "../standards/types";

function create<K extends keyof HTMLElementTagNameMap>(tag: K, props: Partial<HTMLElementTagNameMap[K]> = {}, ...children: Array<Node | string>): HTMLElementTagNameMap[K] {
  const element = document.createElement(tag);
  Object.assign(element, props);
  element.append(...children);
  return element;
}

const errorText = (error: unknown): string => (error instanceof Error ? error.message : String(error));

/**
 * Mục "Biến văn bản này thành mẫu" (issue #2): xem trước các vùng sẽ thay bằng Content Control, chụp bản gốc để khôi
 * phục, rồi tạo mẫu. Lưu mẫu bằng File > Save As > Word Template (*.dotx) hoặc .docx.
 */
export function mountTemplateMaker(container: HTMLElement, deps: { currentStandard: () => StandardProfile }): void {
  let busy = false;
  let snapshot: DocumentSnapshot | null = null;
  const message = create("p", { id: "tm-message" });
  message.setAttribute("role", "status");
  const preview = create("div", { id: "tm-preview" });
  const look = create("button", { type: "button", id: "tm-preview-button", textContent: "Xem các vùng" });
  const make = create("button", { type: "button", id: "tm-make", textContent: "Biến thành mẫu", className: "btn-primary" });
  const restore = create("button", { type: "button", id: "tm-restore", textContent: "Khôi phục bản gốc", hidden: true });
  container.append(
    create("p", { className: "note", textContent: "Giữ nguyên bố cục trang, bảng thể thức và style; thay số, ngày tháng, trích yếu, căn cứ, thân, nơi nhận, người ký bằng ô điền (Content Control) có tiêu đề. Nên làm trên bản sao của văn bản. Xong thì lưu bằng File > Save As > Word Template (*.dotx)." }),
    create("div", { className: "rp-actions" }, look, make, restore),
    preview,
    message,
  );

  function show(kind: "ok" | "warning" | "error" | "", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  function renderPlan(plan: TemplatePlan): void {
    const list = create("ol");
    for (const f of plan.fields) list.append(create("li", {}, create("strong", { textContent: f.title }), f.placeholder === null ? " (giữ nguyên chữ)" : ` → "${f.placeholder}"`));
    preview.replaceChildren(plan.fields.length > 0 ? list : create("p", { className: "note", textContent: "Không nhận ra vùng nào." }), ...plan.notes.map((note) => create("p", { className: "note", textContent: note })));
  }

  async function guarded(task: () => Promise<void>): Promise<void> {
    if (busy) return;
    busy = true;
    look.disabled = make.disabled = restore.disabled = true;
    try {
      await task();
    } catch (error) {
      show("error", `Không làm được: ${errorText(error)}`);
    } finally {
      busy = false;
      look.disabled = make.disabled = restore.disabled = false;
    }
  }

  look.addEventListener("click", () =>
    void guarded(async () => {
      renderPlan(await previewTemplate(deps.currentStandard()));
      show("", "");
    }),
  );

  make.addEventListener("click", () =>
    void guarded(async () => {
      show("ok", "Đang chụp bản gốc để có thể khôi phục…");
      const taken = await takeSnapshot();
      const result = await makeTemplate(deps.currentStandard());
      renderPlan(result.plan);
      if (result.applied === 0) {
        show("warning", "Không nhận ra vùng nào để tạo mẫu; tài liệu giữ nguyên.");
        return;
      }
      snapshot = taken;
      restore.hidden = false;
      show("ok", `Đã tạo ${result.applied} ô điền. Lưu bằng File > Save As > Word Template (*.dotx). Muốn quay lại thì bấm Khôi phục bản gốc hoặc Ctrl+Z.`);
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
}
