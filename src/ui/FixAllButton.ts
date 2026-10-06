import { describeSnapshot, restoreSnapshot, takeSnapshot, type DocumentSnapshot } from "../office/snapshot";
import { buildPreview, planPipeline, runPipeline, summarizeReport, type PlannedStep, type Preview } from "../pipeline/runner";
import { createFixAllRegistry, type FixAllContext } from "../pipeline/steps";
import type { StepRegistry } from "../pipeline/runner";
import type { FormatValues } from "../settings/formatOptions";
import type { StyleSpec } from "../ooxml/styleSpecs";
import type { StandardProfile } from "../standards/types";

export interface FixAllDeps {
  currentStandard: () => StandardProfile;
  values: () => FormatValues;
  styleSpecs: () => StyleSpec[];
  behavior: () => { localFormatting: "keep" | "clear" };
  /** Các task sau truyền registry đã đăng ký thêm bước; bỏ trống thì dùng các bước mặc định. */
  registry?: StepRegistry<FixAllContext>;
  /** Chụp và khôi phục tài liệu; bỏ trống thì dùng bản thật qua Word (truyền vào khi muốn thử giao diện riêng). */
  snapshots?: {
    take: () => Promise<DocumentSnapshot>;
    restore: (snapshot: DocumentSnapshot) => Promise<{ problems: string[] }>;
  };
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

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/** Nút "Chỉnh toàn bộ" (T6C): xem trước, xác nhận, chụp bản gốc, chạy các bước, báo cáo, khôi phục. */
export function mountFixAllButton(container: HTMLElement, deps: FixAllDeps): void {
  const registry = deps.registry ?? createFixAllRegistry();
  const snapshots = deps.snapshots ?? { take: takeSnapshot, restore: restoreSnapshot };
  let snapshot: DocumentSnapshot | null = null;
  let busy = false;

  const main = create("button", { id: "fa-main", type: "button", textContent: "Chỉnh toàn bộ" });
  const restore = create("button", { id: "fa-restore", type: "button", textContent: "Khôi phục bản trước khi chỉnh", hidden: true });
  const panel = create("div", { id: "fa-panel" });
  const message = create("p", { id: "fa-message" });
  message.setAttribute("role", "status");
  container.append(
    create("p", { className: "note", textContent: "Chuẩn hoá cả tài liệu một lần theo chuẩn và Tuỳ chọn nhanh đang chọn: trang, style, phần thân, bảng biểu. Có xem trước và khôi phục." }),
    create("div", {}, main, " ", restore),
    panel,
    message,
  );

  function show(kind: "ok" | "warning" | "error" | "", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  function setBusy(value: boolean): void {
    busy = value;
    main.disabled = value;
    restore.disabled = value;
  }

  function context(): FixAllContext {
    return {
      profile: deps.currentStandard(),
      values: deps.values(),
      specs: deps.styleSpecs(),
      localFormatting: deps.behavior().localFormatting,
    };
  }

  function renderPreview(preview: Preview, planned: PlannedStep[], ctx: FixAllContext): void {
    const box = create("div", { className: "fa-box", id: "fa-preview" }, create("strong", { textContent: preview.heading }));
    const list = create("ul");
    for (const item of preview.items) {
      const entry = create("li", { className: item.willRun ? "" : "fa-skip" }, create("b", { textContent: item.label }));
      for (const line of item.lines) entry.append(create("div", { textContent: line }));
      list.append(entry);
    }
    box.append(list);
    if (preview.nothingToDo) {
      box.append(create("p", { textContent: "Tài liệu đã đúng theo chuẩn và tuỳ chọn hiện tại, không có gì để chỉnh." }));
      box.append(button("Đóng", () => panel.replaceChildren()));
    } else {
      box.append(
        create("p", {
          className: "note",
          textContent: "Trước khi chạy, add-in chụp lại thân văn bản và thiết lập trang để khôi phục. Phần nào không nhận diện chắc chắn thì được bỏ qua, không đoán. Có thể thử Ctrl+Z để hoàn tác; nếu không đủ, dùng nút Khôi phục.",
        }),
        button("Chạy", () => void execute(planned, ctx), { id: "fa-run" }),
        button("Huỷ", () => {
          panel.replaceChildren();
          show("", "");
        }),
      );
    }
    panel.replaceChildren(box);
  }

  function button(text: string, onClick: () => void, props: Partial<HTMLButtonElement> = {}): HTMLButtonElement {
    const element = create("button", { type: "button", textContent: text, ...props });
    element.addEventListener("click", onClick);
    return element;
  }

  async function onMain(): Promise<void> {
    if (busy) return;
    setBusy(true);
    panel.replaceChildren();
    show("ok", "Đang kiểm tra tài liệu…");
    try {
      const ctx = context();
      const planned = await planPipeline(registry.steps(), ctx);
      renderPreview(buildPreview(ctx.profile.name, planned), planned, ctx);
      show("", "");
    } catch (error) {
      show("error", `Không lập được kế hoạch: ${errorText(error)}`);
    } finally {
      setBusy(false);
    }
  }

  async function execute(planned: PlannedStep[], ctx: FixAllContext): Promise<void> {
    if (busy) return;
    setBusy(true);
    panel.replaceChildren();
    try {
      show("ok", "Đang chụp bản trước khi chỉnh…");
      let taken: DocumentSnapshot;
      try {
        taken = await snapshots.take();
      } catch (error) {
        show("error", `Không chụp được bản trước khi chỉnh nên chưa chạy gì, tài liệu giữ nguyên: ${errorText(error)}`);
        return;
      }
      const steps = registry.steps();
      const report = await runPipeline(steps, ctx, planned, (event) => {
        if (event.phase === "start") show("ok", `Đang chạy: ${steps.find((s) => s.id === event.stepId)?.label ?? event.stepId}…`);
      });
      if (report.changed) {
        snapshot = taken;
        restore.hidden = false;
      }
      const summary = summarizeReport(report);
      const box = create("div", { className: "fa-box", id: "fa-report" }, create("strong", { textContent: report.ok ? "Kết quả" : "Chưa hoàn tất" }));
      const list = create("ul");
      for (const line of summary.lines) list.append(create("li", { textContent: line }));
      box.append(list);
      if (!report.ok) {
        box.append(create("p", { textContent: "Các bước trước bước lỗi đã được áp dụng. Bấm Khôi phục để về bản trước khi chỉnh, hoặc sửa nguyên nhân rồi bấm Chỉnh toàn bộ lại (các bước đã đúng sẽ được bỏ qua)." }));
      }
      panel.replaceChildren(box);
      show(summary.kind, report.ok ? (report.changed ? "Đã chỉnh xong." : "Không có gì thay đổi.") : "Có bước bị lỗi, xem chi tiết bên trên.");
    } catch (error) {
      show("error", `Lỗi ngoài dự kiến: ${errorText(error)}. Bấm Khôi phục nếu tài liệu có vẻ đã đổi.`);
      if (snapshot) restore.hidden = false;
    } finally {
      setBusy(false);
    }
  }

  async function doRestore(saved: DocumentSnapshot): Promise<void> {
    if (busy) return;
    setBusy(true);
    panel.replaceChildren();
    try {
      show("ok", "Đang khôi phục…");
      const result = await snapshots.restore(saved);
      if (result.problems.length === 0) {
        snapshot = null;
        restore.hidden = true;
        show("ok", `Đã khôi phục về ${describeSnapshot(saved)}. Style ND30_*/HD05_* giữ giá trị mới.`);
      } else {
        show("warning", `Đã khôi phục nhưng còn điểm chưa khớp: ${result.problems.join("; ")}.`);
      }
    } catch (error) {
      show("error", `Không khôi phục được: ${errorText(error)}`);
    } finally {
      setBusy(false);
    }
  }

  restore.addEventListener("click", () => {
    if (busy || !snapshot) return;
    const saved = snapshot;
    const box = create("div", { className: "fa-box", id: "fa-restore-confirm" });
    box.setAttribute("role", "alert");
    box.append(
      create("p", {
        textContent: `Khôi phục sẽ thay lại toàn bộ thân văn bản và thiết lập trang bằng ${describeSnapshot(saved)}. Mọi chỉnh sửa của bạn sau thời điểm đó sẽ mất; bình luận hoặc đánh dấu có thể bị ảnh hưởng.`,
      }),
      button("Khôi phục", () => void doRestore(saved)),
      button("Huỷ", () => panel.replaceChildren()),
    );
    panel.replaceChildren(box);
  });

  main.addEventListener("click", () => {
    void onMain();
  });
}
