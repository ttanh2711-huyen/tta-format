import { mountApp, setStatus } from "./app";

/**
 * Trang xem trước giao diện (chỉ có ở bản phát triển, không vào bản phát hành): giả lập Office và Word tối thiểu để dựng
 * bảng TTA-Format ngoài Word, dùng để xem và chụp giao diện. Mọi thao tác ghi vào tài liệu báo lỗi "chỉ xem trước".
 */
type Loose = Record<string, unknown>;

const reject = (): Promise<never> => Promise.reject(new Error("Đang xem trước giao diện, không có tài liệu Word."));

const globals = globalThis as unknown as Loose;
globals.Office = {
  HostType: { Word: "Word" },
  onReady: (callback: (info: { host: string }) => void) => Promise.resolve(callback({ host: "Word" })),
  context: { requirements: { isSetSupported: () => true }, document: {} },
};
globals.Word = { run: reject };

mountApp();
setStatus("Xem trước giao diện (không có Word).");
