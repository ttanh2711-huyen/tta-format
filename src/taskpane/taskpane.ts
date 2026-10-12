import { mountApp, setStatus } from "./app";

/** Lỗi không ai bắt: hiện lên dòng trạng thái thay vì im lặng. */
function describeError(reason: unknown): string {
  return reason instanceof Error ? reason.message : String(reason);
}
window.addEventListener("error", (event) => setStatus(`Có lỗi ngoài dự kiến: ${event.message}. Hãy tải lại add-in; nếu còn lỗi, báo kèm nội dung này.`));
window.addEventListener("unhandledrejection", (event) => setStatus(`Có thao tác không hoàn tất: ${describeError(event.reason)}. Tài liệu có thể chưa thay đổi; hãy kiểm tra lại.`));

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus("Add-in này chỉ chạy trong Microsoft Word.");
    return;
  }
  try {
    mountApp();
  } catch (error) {
    setStatus(`Không dựng được giao diện: ${describeError(error)}. Hãy tải lại add-in.`);
  }
});
