import { summarizeIssues } from "../checker/engine";
import { buildReportDocx, toBase64 } from "../checker/exportDocx";
import { formatReport } from "../checker/exportReport";
import { planFixes, type FixPlan } from "../checker/fixAll";
import type { Issue } from "../checker/types";
import { applyFixes } from "../office/applyFix";
import { checkDocument, runFixAll } from "../office/fixAll";
import { isReadDocumentSupported, selectParagraph } from "../office/readDocument";
import { describeSnapshot, restoreSnapshot, takeSnapshot, type DocumentSnapshot } from "../office/snapshot";
import { planPipeline, runPipeline, summarizeReport } from "../pipeline/runner";
import { createFixAllRegistry, type FixAllContext } from "../pipeline/steps";
import type { StandardProfile } from "../standards/types";

export interface ReportPanelDeps {
  currentStandard: () => StandardProfile;
  /** Ngữ cảnh của "Chỉnh toàn bộ" (tuỳ chọn nhanh, style); có thì hiện nút "Sửa nhanh". */
  pipelineContext?: () => FixAllContext;
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

/**
 * Bảng "Kiểm tra văn bản" (T13, T15): đọc tài liệu, đối chiếu quy tắc của chuẩn đang chọn, liệt kê lỗi, nhảy tới chỗ lỗi,
 * sửa từng lỗi, "Sửa tất cả" có xem trước và hoàn tác, xuất báo cáo.
 */
export function mountReportPanel(container: HTMLElement, deps: ReportPanelDeps): { refresh: () => void } {
  let busy = false;
  const snapshots = deps.snapshots ?? { take: takeSnapshot, restore: restoreSnapshot };
  let undoSnapshot: DocumentSnapshot | null = null;
  let issuesNow: Issue[] = [];

  const check = create("button", { id: "rp-check", type: "button", textContent: "Kiểm tra", title: "Đọc văn bản và liệt kê lỗi (chỉ đọc)" });
  const summary = create("p", { id: "rp-summary" });
  summary.setAttribute("role", "status");
  const fixAll = create("button", { id: "rp-fix-all", type: "button", textContent: "Sửa tất cả", hidden: true });
  const undo = create("button", { id: "rp-undo", type: "button", textContent: "Hoàn tác", title: "Hoàn tác lần Sửa tất cả hoặc Sửa nhanh gần nhất", hidden: true });
  const quick = create("button", { id: "rp-quick", type: "button", textContent: "Sửa nhanh", hidden: !deps.pipelineContext, title: "Chỉnh toàn bộ rồi sửa tất cả lỗi sửa được, một lần bấm; hoàn tác được." });
  const exportButton = create("button", { id: "rp-export", type: "button", textContent: "Báo cáo", title: "Xuất báo cáo kiểm tra (sao chép, .txt, .docx)", hidden: true });
  const preview = create("div", { id: "rp-preview", className: "fa-box", hidden: true });
  const exported = create("div", { id: "rp-exported", hidden: true });
  const list = create("ol", { id: "rp-list" });
  container.append(
    create("h2", { textContent: "Kiểm tra văn bản" }),
    create("p", { className: "note", textContent: "Đối chiếu trang, phông, cỡ chữ, căn lề, thụt dòng, giãn đoạn, giãn dòng, khối đầu và khối cuối của văn bản với chuẩn đang chọn. Kiểm tra chỉ đọc; nút \"Sửa\" sửa đúng chỗ của từng lỗi, \"Sửa tất cả\" cho xem trước rồi mới sửa." }),
    create("div", { className: "rp-actions" }, check, quick, fixAll, undo, exportButton),
    summary,
    preview,
    exported,
    list,
  );

  function show(kind: "ok" | "warning" | "error" | "", text: string): void {
    summary.className = kind;
    summary.textContent = text;
  }

  function renderIssue(issue: Issue): HTMLElement {
    const where = issue.location.kind === "page" ? `Trang, phần ${issue.location.section}` : `Đoạn: "${issue.location.excerpt}"`;
    const item = create("li", { className: issue.severity });
    item.append(
      create("strong", { textContent: issue.severity === "error" ? "Lỗi: " : "Cảnh báo: " }),
      issue.message,
      create("div", { className: "note", textContent: `${where} · ${issue.reference}` }),
    );
    const actions = create("div", { className: "rp-actions" });
    if (issue.location.kind === "paragraph") {
      const index = issue.location.index;
      const go = create("button", { type: "button", textContent: "Đến vị trí", className: "rp-go" });
      go.addEventListener("click", () => {
        void selectParagraph(index).then((found) => {
          if (!found) show("warning", "Không tìm thấy đoạn này nữa; hãy kiểm tra lại.");
        });
      });
      actions.append(go);
    }
    const fix = issue.fix;
    if (fix) {
      const button = create("button", { type: "button", textContent: "Sửa", className: "rp-fix", title: fix.label });
      button.addEventListener("click", () => {
        button.disabled = true;
        void applyFixes([fix])
          .then((result) => {
            if (result.skipped.length > 0) show("warning", result.skipped.join("; "));
            return onCheck();
          })
          .catch((error: unknown) => show("error", `Không sửa được: ${errorText(error)}`));
      });
      actions.append(button);
    }
    if (actions.children.length > 0) item.append(actions);
    return item;
  }

  function planNow(): FixPlan {
    return planFixes(issuesNow);
  }

  function closePreview(): void {
    preview.hidden = true;
    preview.replaceChildren();
  }

  /** Hộp xem trước: liệt kê từng thay đổi sẽ làm; chỉ khi bấm "Áp dụng" mới sửa. */
  function openPreview(): void {
    const plan = planNow();
    if (plan.fixes.length === 0) return;
    const items = plan.changes.map((change) => create("li", {}, `${change.where}: ${change.label}`));
    const apply = create("button", { type: "button", textContent: `Áp dụng ${plan.fixes.length} thay đổi`, id: "rp-apply" });
    const cancel = create("button", { type: "button", textContent: "Huỷ", id: "rp-cancel" });
    cancel.addEventListener("click", closePreview);
    apply.addEventListener("click", () => void onApplyAll());
    const rest = plan.manual.length > 0 ? `; ${plan.manual.length} lỗi còn lại cần bạn xem` : "";
    preview.replaceChildren(
      create("p", {
        textContent: `Sẽ sửa ${plan.fixes.length} chỗ${rest}. Sau khi sửa, add-in kiểm tra lại và sửa tiếp nếu còn lỗi sửa được (tối đa vài lượt). Hoàn tác ngay sau đó bằng nút "Hoàn tác" hoặc Ctrl+Z.`,
      }),
      create("ol", {}, ...items),
      create("div", { className: "rp-actions" }, apply, cancel),
    );
    preview.hidden = false;
  }

  async function onApplyAll(): Promise<void> {
    if (busy) return;
    closePreview();
    busy = true;
    fixAll.disabled = true;
    check.disabled = true;
    show("ok", "Đang chụp tài liệu để có thể hoàn tác…");
    let outcome: { kind: "ok" | "warning"; text: string } | undefined;
    try {
      const taken = await snapshots.take();
      show("ok", "Đang sửa…");
      const result = await runFixAll(deps.currentStandard());
      undoSnapshot = taken;
      undo.hidden = false;
      const fixable = result.remaining.filter((issue) => issue.fix).length;
      const note = result.skipped.length > 0 ? ` Không sửa được: ${result.skipped.join("; ")}.` : "";
      outcome = {
        kind: result.skipped.length > 0 || fixable > 0 ? "warning" : "ok",
        text: `Đã sửa ${result.applied} chỗ sau ${result.rounds} lượt; còn ${result.remaining.length} lỗi hoặc cảnh báo (${fixable} sửa được).${note}`,
      };
    } catch (error) {
      show("error", `Không sửa được: ${errorText(error)}`);
    } finally {
      busy = false;
      fixAll.disabled = false;
      check.disabled = false;
    }
    if (outcome) {
      await onCheck();
      show(outcome.kind, outcome.text);
    }
  }

  fixAll.addEventListener("click", openPreview);

  /** Sửa nhanh: một lần chụp, chạy các bước Chỉnh toàn bộ rồi Sửa tất cả; Hoàn tác về bản chụp. */
  async function onQuick(): Promise<void> {
    if (busy || !deps.pipelineContext) return;
    busy = true;
    quick.disabled = fixAll.disabled = check.disabled = true;
    show("ok", "Đang chụp tài liệu để có thể hoàn tác…");
    let outcome: { kind: "ok" | "warning" | "error"; text: string } | undefined;
    try {
      const taken = await snapshots.take();
      const ctx = deps.pipelineContext();
      const registry = createFixAllRegistry();
      const steps = registry.steps();
      show("ok", "Đang chỉnh toàn bộ…");
      const report = await runPipeline(steps, ctx, await planPipeline(steps, ctx));
      if (!report.ok) {
        undoSnapshot = taken;
        undo.hidden = false;
        outcome = { kind: "error", text: `Chỉnh toàn bộ chưa hoàn tất: ${summarizeReport(report).lines.join("; ")}. Có thể hoàn tác.` };
      } else {
        show("ok", "Đang sửa các lỗi còn lại…");
        const result = await runFixAll(ctx.profile);
        undoSnapshot = taken;
        undo.hidden = false;
        const fixable = result.remaining.filter((issue) => issue.fix).length;
        const note = result.skipped.length > 0 ? ` Không sửa được: ${result.skipped.join("; ")}.` : "";
        outcome = {
          kind: result.skipped.length > 0 || fixable > 0 ? "warning" : "ok",
          text: `Sửa nhanh xong: ${result.applied} chỗ sửa thêm sau Chỉnh toàn bộ; còn ${result.remaining.length} lỗi hoặc cảnh báo cần bạn xem.${note}`,
        };
      }
    } catch (error) {
      show("error", `Không sửa nhanh được: ${errorText(error)}`);
    } finally {
      busy = false;
      quick.disabled = fixAll.disabled = check.disabled = false;
    }
    if (outcome) {
      await onCheck();
      show(outcome.kind, outcome.text);
    }
  }

  quick.addEventListener("click", () => void onQuick());

  async function restoreAndRecheck(saved: DocumentSnapshot): Promise<void> {
    let text: { kind: "ok" | "warning"; text: string } | undefined;
    try {
      const result = await snapshots.restore(saved);
      undoSnapshot = null;
      undo.hidden = true;
      text =
        result.problems.length > 0
          ? { kind: "warning", text: `Đã khôi phục về ${describeSnapshot(saved)} nhưng: ${result.problems.join("; ")}.` }
          : { kind: "ok", text: `Đã khôi phục về ${describeSnapshot(saved)}.` };
    } catch (error) {
      show("error", `Không khôi phục được: ${errorText(error)}`);
    } finally {
      busy = false;
      undo.disabled = false;
    }
    if (text) {
      await onCheck();
      show(text.kind, text.text);
    }
  }

  undo.addEventListener("click", () => {
    if (busy || !undoSnapshot) return;
    busy = true;
    undo.disabled = true;
    show("ok", "Đang khôi phục…");
    void restoreAndRecheck(undoSnapshot);
  });

  exportButton.addEventListener("click", () => {
    const text = formatReport(issuesNow, { standardName: deps.currentStandard().name, checkedAt: new Date() });
    const area = create("textarea", { id: "rp-report-text", readOnly: true, value: text, rows: 10 });
    const copy = create("button", { type: "button", textContent: "Sao chép" });
    copy.addEventListener("click", () => {
      area.select();
      void navigator.clipboard?.writeText(text).catch(() => document.execCommand("copy"));
    });
    const file = new Blob(["﻿", text], { type: "text/plain;charset=utf-8" });
    const link = create("a", { textContent: "Tải về .txt", href: URL.createObjectURL(file), download: "bao-cao-kiem-tra.txt" });
    const docx = create("button", { type: "button", textContent: "Mở báo cáo trong Word (.docx)", id: "rp-docx" });
    docx.addEventListener("click", () => {
      docx.disabled = true;
      const base64 = toBase64(buildReportDocx(issuesNow, { standardName: deps.currentStandard().name, checkedAt: new Date() }));
      void Word.run(async (context) => {
        context.application.createDocument(base64).open();
        await context.sync();
      })
        .catch((error: unknown) => show("error", `Không mở được báo cáo .docx: ${errorText(error)}. Dùng bản .txt.`))
        .finally(() => {
          docx.disabled = false;
        });
    });
    exported.replaceChildren(area, create("div", { className: "rp-actions" }, copy, docx, link));
    exported.hidden = false;
  });

  async function onCheck(): Promise<void> {
    if (busy) return;
    busy = true;
    check.disabled = true;
    list.replaceChildren();
    fixAll.hidden = true;
    exportButton.hidden = true;
    closePreview();
    exported.hidden = true;
    show("ok", "Đang đọc tài liệu…");
    try {
      const profile = deps.currentStandard();
      const issues = await checkDocument(profile);
      issuesNow = issues;
      const counts = summarizeIssues(issues);
      for (const issue of issues) list.append(renderIssue(issue));
      const pending = planNow().fixes.length;
      fixAll.hidden = pending === 0;
      fixAll.textContent = `Sửa tất cả (${pending})`;
      fixAll.title = "Xem trước các thay đổi rồi mới áp dụng.";
      exportButton.hidden = false;
      if (issues.length === 0) show("ok", `Không thấy lỗi theo chuẩn ${profile.name}.`);
      else show(counts.errors > 0 ? "error" : "warning", `Chuẩn ${profile.name}: ${counts.errors} lỗi, ${counts.warnings} cảnh báo.`);
    } catch (error) {
      show("error", `Không kiểm tra được: ${errorText(error)}`);
    } finally {
      busy = false;
      check.disabled = false;
    }
  }

  if (isReadDocumentSupported()) {
    check.addEventListener("click", () => void onCheck());
  } else {
    check.disabled = quick.disabled = true;
    show("error", "Kiểm tra văn bản cần Microsoft 365 bản 2507 trở lên (WordApiDesktop 1.3).");
  }

  return {
    refresh: () => {
      list.replaceChildren();
      fixAll.hidden = true;
      exportButton.hidden = true;
      closePreview();
      exported.hidden = true;
      undoSnapshot = null;
      undo.hidden = true;
      if (isReadDocumentSupported()) show("", "");
    },
  };
}
