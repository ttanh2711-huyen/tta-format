import { runChecks, summarizeIssues } from "../checker/engine";
import { headerFooterRules } from "../checker/ruleSets";
import type { Issue } from "../checker/types";
import { applyFixes } from "../office/applyFix";
import { isReadDocumentSupported, readCheckDocument, selectParagraph } from "../office/readDocument";
import type { StandardProfile } from "../standards/types";

export interface ReportPanelDeps {
  currentStandard: () => StandardProfile;
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

/** Bảng "Kiểm tra văn bản" (T13): đọc tài liệu, đối chiếu quy tắc của chuẩn đang chọn, liệt kê lỗi và nhảy tới chỗ lỗi. */
export function mountReportPanel(container: HTMLElement, deps: ReportPanelDeps): { refresh: () => void } {
  let busy = false;
  const check = create("button", { id: "rp-check", type: "button", textContent: "Kiểm tra văn bản" });
  const summary = create("p", { id: "rp-summary" });
  summary.setAttribute("role", "status");
  const list = create("ol", { id: "rp-list" });
  container.append(
    create("h2", { textContent: "Kiểm tra văn bản" }),
    create("p", { className: "note", textContent: "Đối chiếu trang, phông, cỡ chữ, căn lề, thụt dòng, giãn đoạn, giãn dòng, khối đầu và khối cuối của văn bản với chuẩn đang chọn. Kiểm tra chỉ đọc; chỉ nút \"Sửa\" của từng lỗi mới sửa đúng chỗ đó." }),
    check,
    summary,
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
    if (issue.location.kind === "paragraph") {
      const index = issue.location.index;
      const go = create("button", { type: "button", textContent: "Đến vị trí", className: "rp-go" });
      go.addEventListener("click", () => {
        void selectParagraph(index).then((found) => {
          if (!found) show("warning", "Không tìm thấy đoạn này nữa; hãy kiểm tra lại.");
        });
      });
      item.append(go);
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
      item.append(button);
    }
    return item;
  }

  async function onCheck(): Promise<void> {
    if (busy) return;
    busy = true;
    check.disabled = true;
    list.replaceChildren();
    show("ok", "Đang đọc tài liệu…");
    try {
      const profile = deps.currentStandard();
      const issues = runChecks(await readCheckDocument(), profile, headerFooterRules(profile));
      const counts = summarizeIssues(issues);
      for (const issue of issues) list.append(renderIssue(issue));
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
    check.disabled = true;
    show("error", "Kiểm tra văn bản cần Microsoft 365 bản 2507 trở lên (WordApiDesktop 1.3).");
  }

  return {
    refresh: () => {
      list.replaceChildren();
      if (isReadDocumentSupported()) show("", "");
    },
  };
}
