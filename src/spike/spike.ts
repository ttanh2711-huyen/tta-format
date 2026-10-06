import { PROBES } from "./probes";
import { goNoGo, summarize, toMarkdown, type ProbeResult, type SpikeReport } from "./report";
import { collectEnvironment, runAll, runProbe } from "./runner";

const results = new Map<string, ProbeResult>();

function element<T extends HTMLElement>(id: string): T {
  const found = document.getElementById(id);
  if (!found) {
    throw new Error(`Thiếu phần tử #${id}`);
  }
  return found as T;
}

function buildReport(running?: string): SpikeReport {
  return {
    generatedAt: new Date().toISOString(),
    environment: collectEnvironment(),
    results: PROBES.map((p) => results.get(p.id)).filter((r): r is ProbeResult => r !== undefined),
    ...(running ? { running } : {}),
  };
}

function renderRow(probeId: string): void {
  const probe = PROBES.find((p) => p.id === probeId);
  const row = document.getElementById(`row-${probeId}`);
  if (!probe || !row) return;
  const result = results.get(probeId);
  row.querySelector<HTMLElement>(".badge")!.textContent = result ? result.status : "chưa chạy";
  row.querySelector<HTMLElement>(".badge")!.className = `badge ${result?.status ?? ""}`;
  row.querySelector<HTMLElement>(".detail")!.textContent = result ? `${result.detail} (${result.durationMs} ms)` : "";
}

function renderSummary(): void {
  const all = PROBES.map((p) => results.get(p.id)).filter((r): r is ProbeResult => r !== undefined);
  const counts = summarize(all);
  const { verdict } = goNoGo(all);
  element("summary").textContent =
    `pass ${counts.pass}, fail ${counts.fail}, không hỗ trợ ${counts.unsupported}, thủ công ${counts.manual}, bỏ qua ${counts.skipped}. Kết luận tạm: ${verdict}`;
}

function setBusy(busy: boolean): void {
  const confirmed = element<HTMLInputElement>("confirm").checked;
  element<HTMLButtonElement>("btn-all").disabled = busy || !confirmed;
  document.querySelectorAll<HTMLButtonElement>(".run-one").forEach((b) => {
    b.disabled = busy || !confirmed;
  });
  element<HTMLButtonElement>("btn-copy").disabled = busy || results.size === 0;
}

function record(result: ProbeResult): void {
  results.set(result.id, result);
  renderRow(result.id);
  renderSummary();
}

interface SaveOptions {
  /** Mã phép thử sắp chạy; lưu trước để nếu Word đóng đột ngột còn biết là mục nào. */
  running?: string;
  /** "spike-rieng" cho lượt chạy riêng, để không đè báo cáo đầy đủ. */
  name?: "spike-results" | "spike-rieng";
  quiet?: boolean;
}

async function saveReport(options: SaveOptions = {}): Promise<void> {
  const report = buildReport(options.running);
  const markdown = toMarkdown(report);
  const file = options.name ?? "spike-results";
  const status = element("save-status");
  try {
    const response = await fetch("/spike-report", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ report, markdown, name: file }),
    });
    if (!options.quiet) {
      status.textContent = response.ok
        ? `Đã lưu báo cáo vào docs/${file}.json và docs/${file}.md (máy chủ dev).`
        : `Máy chủ không nhận báo cáo (HTTP ${response.status}). Dùng nút "Sao chép báo cáo".`;
    }
  } catch {
    status.textContent = 'Không gửi được báo cáo tới máy chủ dev. Dùng nút "Sao chép báo cáo".';
  }
}

async function onRunAll(): Promise<void> {
  setBusy(true);
  element("save-status").textContent = "";
  try {
    await runAll((probe) => saveReport({ running: probe.id, quiet: true }), record);
    await saveReport();
  } finally {
    setBusy(false);
  }
}

async function onRunOne(probeId: string): Promise<void> {
  const probe = PROBES.find((p) => p.id === probeId);
  if (!probe) return;
  setBusy(true);
  try {
    await saveReport({ running: probe.id, name: "spike-rieng", quiet: true });
    record(await runProbe(probe));
    await saveReport({ name: "spike-rieng" });
  } finally {
    setBusy(false);
  }
}

async function onCopy(): Promise<void> {
  const markdown = toMarkdown(buildReport());
  try {
    await navigator.clipboard.writeText(markdown);
    element("save-status").textContent = "Đã sao chép báo cáo (Markdown) vào clipboard.";
  } catch {
    element("save-status").textContent = "Không sao chép được vào clipboard.";
  }
}

function buildList(): void {
  const list = element("list");
  for (const probe of PROBES) {
    const row = document.createElement("div");
    row.className = "row";
    row.id = `row-${probe.id}`;

    const head = document.createElement("div");
    head.className = "head";
    const title = document.createElement("span");
    title.className = "title";
    title.textContent = `${probe.id}${probe.critical ? " *" : ""}${probe.risky ? " ⚠" : ""}  ${probe.title}`;
    const badge = document.createElement("span");
    badge.className = "badge";
    badge.textContent = "chưa chạy";
    head.append(title, badge);

    const detail = document.createElement("div");
    detail.className = "detail";

    const run = document.createElement("button");
    run.type = "button";
    run.className = "run-one";
    run.textContent = "Chạy riêng";
    run.disabled = true;
    run.addEventListener("click", () => void onRunOne(probe.id));

    row.append(head, detail, run);
    list.append(row);
  }
}

Office.onReady((info) => {
  buildList();
  if (info.host !== Office.HostType.Word) {
    element("summary").textContent = "Trang này chỉ chạy trong Microsoft Word.";
    return;
  }
  element<HTMLInputElement>("confirm").addEventListener("change", () => setBusy(false));
  element("btn-all").addEventListener("click", () => void onRunAll());
  element("btn-copy").addEventListener("click", () => void onCopy());
  element("btn-back").addEventListener("click", () => {
    window.location.href = "taskpane.html";
  });
  element("summary").textContent = "Chưa chạy. Dấu * là mục bắt buộc cho quyết định go/no-go; dấu ⚠ là mục rủi ro, chỉ chạy riêng.";
});
