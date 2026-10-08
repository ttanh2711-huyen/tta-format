import { insertSampleParagraph } from "../office/insertSample";
import { isDocumentStandardSupported, readDocumentSnapshot, writeDocumentNoAsk, writeDocumentStandard } from "../office/documentStandard";
import { browserStore } from "../settings/storage";
import { mountFixAllButton } from "../ui/FixAllButton";
import { mountHeaderForm } from "../ui/HeaderForm";
import { mountFooterForm } from "../ui/FooterForm";
import { mountFormatPanel } from "../ui/FormatPanel";
import { mountPageSetupPanel } from "../ui/PageSetupPanel";
import { mountQuickOptions } from "../ui/QuickOptions";
import { mountStandardSwitcher } from "../ui/StandardSwitcher";

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

  const switcherEl = document.getElementById("standard-switcher");
  const fixAllEl = document.getElementById("fix-all");
  const quickEl = document.getElementById("quick-options");
  const pageEl = document.getElementById("page-setup-panel");
  const formatEl = document.getElementById("format-panel");
  const headerEl = document.getElementById("header-form");
  const footerEl = document.getElementById("footer-form");
  if (!switcherEl || !fixAllEl || !quickEl || !pageEl || !formatEl || !headerEl || !footerEl) {
    setStatus("Thiếu vùng hiển thị trong trang; hãy tải lại add-in.");
    return;
  }

  const store = browserStore();
  const documentSupported = isDocumentStandardSupported();
  const switcher = mountStandardSwitcher(switcherEl, {
    store,
    readSnapshot: documentSupported ? readDocumentSnapshot : undefined,
    writeStandard: documentSupported ? writeDocumentStandard : undefined,
    writeNoAsk: documentSupported ? writeDocumentNoAsk : undefined,
  });
  const quick = mountQuickOptions(quickEl, { currentStandard: switcher.currentStandard, store });
  const header = mountHeaderForm(headerEl, { currentStandard: switcher.currentStandard, values: quick.values, store });
  const footer = mountFooterForm(footerEl, { currentStandard: switcher.currentStandard, values: quick.values });
  switcher.onChange(() => {
    quick.refresh();
    header.refresh();
    footer.refresh();
  });
  mountFixAllButton(fixAllEl, { currentStandard: switcher.currentStandard, values: quick.values, styleSpecs: quick.styleSpecs, behavior: quick.behavior });
  mountPageSetupPanel(pageEl, { currentStandard: switcher.currentStandard, values: quick.values, styleSpecs: quick.styleSpecs });
  mountFormatPanel(formatEl, { currentStandard: switcher.currentStandard, behavior: quick.behavior, styleSpecs: quick.styleSpecs });
  setStatus("Sẵn sàng.");
  void switcher.start();
});
