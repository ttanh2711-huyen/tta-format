import { insertSampleParagraph } from "../office/insertSample";
import { mountFormatPanel } from "../ui/FormatPanel";
import { mountPageSetupPanel } from "../ui/PageSetupPanel";

function setStatus(message: string): void {
  const status = document.getElementById("status");
  if (status) {
    status.textContent = message;
  }
}

async function onInsertSample(): Promise<void> {
  try {
    await Word.run(async (context) => {
      insertSampleParagraph(context.document.body);
      await context.sync();
    });
    setStatus("Đã chèn đoạn văn bản thử nghiệm vào cuối tài liệu.");
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    setStatus(`Không chèn được văn bản: ${reason}`);
  }
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus("Add-in này chỉ chạy trong Microsoft Word.");
    return;
  }
  document.getElementById("btn-sample")?.addEventListener("click", () => {
    void onInsertSample();
  });
  const panel = document.getElementById("page-setup-panel");
  const formatPanel = document.getElementById("format-panel");
  if (panel) {
    const { currentStandard } = mountPageSetupPanel(panel);
    if (formatPanel) {
      mountFormatPanel(formatPanel, currentStandard);
    }
  }
  setStatus("Sẵn sàng.");
});
