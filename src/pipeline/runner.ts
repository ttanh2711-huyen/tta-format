/**
 * Khung chạy nhiều bước cho nút "Chỉnh toàn bộ" (T6C). Mỗi bước tự khai báo kế hoạch (để hiện hộp xem
 * trước) và cách chạy; các task sau chỉ cần đăng ký thêm bước vào `StepRegistry`, không sửa khung này.
 */

export interface StepPlan {
  /** Bước có việc để làm không; false thì bị bỏ qua kèm `skipReason`. */
  willRun: boolean;
  /** Các dòng mô tả việc sẽ làm, hiện trong hộp xem trước. */
  lines: string[];
  skipReason?: string;
}

export interface StepResult {
  status: "done" | "skipped";
  kind: "ok" | "warning";
  text: string;
}

export interface PipelineStep<Ctx> {
  id: string;
  label: string;
  /** Thứ tự chạy, nhỏ chạy trước. */
  order: number;
  /** Chỉ đọc tài liệu, không được sửa. */
  plan(ctx: Ctx): Promise<StepPlan>;
  run(ctx: Ctx): Promise<StepResult>;
}

export class StepRegistry<Ctx> {
  private readonly items: Array<PipelineStep<Ctx>> = [];

  register(step: PipelineStep<Ctx>): void {
    if (this.items.some((item) => item.id === step.id)) {
      throw new Error(`Bước "${step.id}" đã được đăng ký.`);
    }
    this.items.push(step);
  }

  /** Các bước theo thứ tự chạy; cùng `order` thì theo thứ tự đăng ký. */
  steps(): Array<PipelineStep<Ctx>> {
    return this.items
      .map((step, index) => ({ step, index }))
      .sort((a, b) => a.step.order - b.step.order || a.index - b.index)
      .map(({ step }) => step);
  }
}

export interface PlannedStep extends StepPlan {
  id: string;
  label: string;
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export async function planPipeline<Ctx>(steps: ReadonlyArray<PipelineStep<Ctx>>, ctx: Ctx): Promise<PlannedStep[]> {
  const planned: PlannedStep[] = [];
  for (const step of steps) {
    try {
      planned.push({ id: step.id, label: step.label, ...(await step.plan(ctx)) });
    } catch (error) {
      planned.push({ id: step.id, label: step.label, willRun: false, lines: [], skipReason: `không đọc được tài liệu để lập kế hoạch (${messageOf(error)})` });
    }
  }
  return planned;
}

export type StepState = "done" | "skipped" | "failed" | "notRun";

export interface StepReport {
  id: string;
  label: string;
  state: StepState;
  kind: "ok" | "warning" | "error";
  text: string;
}

export interface PipelineReport {
  steps: StepReport[];
  ok: boolean;
  /** Có bước nào đã sửa tài liệu (kể cả khi sau đó có bước lỗi). */
  changed: boolean;
  failedStepId?: string;
}

export interface ProgressEvent {
  phase: "start" | "end";
  stepId: string;
}

/**
 * Chạy các bước theo thứ tự. Bước lỗi thì dừng: các bước trước giữ nguyên kết quả đã có, các bước sau
 * ghi là chưa chạy, để báo cáo nói rõ tài liệu đang ở đâu.
 */
export async function runPipeline<Ctx>(
  steps: ReadonlyArray<PipelineStep<Ctx>>,
  ctx: Ctx,
  plan?: readonly PlannedStep[],
  onProgress?: (event: ProgressEvent) => void,
): Promise<PipelineReport> {
  const planned = plan ?? (await planPipeline(steps, ctx));
  const reports: StepReport[] = [];
  let failedStepId: string | undefined;
  let changed = false;

  for (const step of steps) {
    if (failedStepId !== undefined) {
      reports.push({ id: step.id, label: step.label, state: "notRun", kind: "warning", text: "chưa chạy" });
      continue;
    }
    const entry = planned.find((p) => p.id === step.id);
    if (entry && !entry.willRun) {
      reports.push({ id: step.id, label: step.label, state: "skipped", kind: "ok", text: entry.skipReason ?? "không có việc để làm" });
      continue;
    }
    onProgress?.({ phase: "start", stepId: step.id });
    try {
      const result = await step.run(ctx);
      reports.push({ id: step.id, label: step.label, state: result.status, kind: result.kind, text: result.text });
      if (result.status === "done") changed = true;
    } catch (error) {
      failedStepId = step.id;
      reports.push({ id: step.id, label: step.label, state: "failed", kind: "error", text: messageOf(error) });
    }
    onProgress?.({ phase: "end", stepId: step.id });
  }

  return { steps: reports, ok: failedStepId === undefined, changed, failedStepId };
}

export interface PreviewItem {
  label: string;
  willRun: boolean;
  lines: string[];
}

export interface Preview {
  heading: string;
  items: PreviewItem[];
  canRun: boolean;
  nothingToDo: boolean;
}

export function buildPreview(standardName: string, planned: readonly PlannedStep[]): Preview {
  const items = planned.map((step) => ({
    label: step.label,
    willRun: step.willRun,
    lines: step.willRun ? step.lines : [`Bỏ qua: ${step.skipReason ?? "không có việc để làm"}`],
  }));
  const canRun = planned.some((step) => step.willRun);
  return { heading: `Chỉnh toàn bộ theo chuẩn ${standardName}`, items, canRun, nothingToDo: !canRun };
}

export function summarizeReport(report: PipelineReport): { kind: "ok" | "warning" | "error"; lines: string[] } {
  const lines = report.steps.map((step) => {
    if (step.state === "failed") return `${step.label}: LỖI: ${step.text}`;
    if (step.state === "notRun") return `${step.label}: chưa chạy`;
    if (step.state === "skipped") return `${step.label}: bỏ qua, ${step.text}`;
    return `${step.label}: ${step.text}`;
  });
  const kind = !report.ok ? "error" : report.steps.some((step) => step.kind === "warning") ? "warning" : "ok";
  return { kind, lines };
}
