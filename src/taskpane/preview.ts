import type { Issue } from "../checker/types";
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

/** Lỗi mẫu cho nút "Kiểm tra" của bản xem trước, để xem cách gom nhóm. */
const at = (index: number, excerpt: string) => ({ kind: "paragraph" as const, index, excerpt });
const SAMPLE_ISSUES: Issue[] = [
  { ruleId: "page.margin.left", severity: "error", message: "Lề trái 25 mm không đúng; chuẩn cho phép 30-35 mm.", reference: "Phụ lục I, mục I.5", location: { kind: "page", section: 1 } },
  { ruleId: "font.size", severity: "error", message: "Cỡ chữ 12 pt ngoài quy định cho nội dung: 13-14 pt.", reference: "Phụ lục I", location: at(7, "Thực hiện Kế hoạch số 12/KH-UBND…"), fix: { kind: "font", index: 7, sizePt: 14, label: "Cỡ 14 pt" } },
  { ruleId: "font.size", severity: "error", message: "Cỡ chữ 12 pt ngoài quy định cho nội dung: 13-14 pt.", reference: "Phụ lục I", location: at(9, "Đề nghị các phòng, ban…"), fix: { kind: "font", index: 9, sizePt: 14, label: "Cỡ 14 pt" } },
  { ruleId: "paragraph.firstLineIndent", severity: "error", message: "Dòng đầu của đoạn chưa thụt vào; cần 10 mm.", reference: "Phụ lục I", location: at(8, "Ủy ban nhân dân xã báo cáo…"), fix: { kind: "paragraph", index: 8, firstLineIndentMm: 10, label: "Thụt 10 mm" } },
  { ruleId: "paragraph.spacing", severity: "warning", message: "Giãn đoạn trước và sau tổng 0 pt; tối thiểu 6 pt.", reference: "Phụ lục I", location: at(10, "Trân trọng./.") },
  { ruleId: "header.date.italic", severity: "error", message: "Địa danh và ngày tháng phải in nghiêng.", reference: "Phụ lục I", location: at(3, "Cần Thơ, ngày 16 tháng 10 năm 2026"), fix: { kind: "font", index: 3, italic: true, label: "In nghiêng" } },
  { ruleId: "cap.referenceLower", severity: "warning", message: "Khi viện dẫn, \"khoản\", \"điểm\" viết thường: \"Khoản 2\".", reference: "Phụ lục II", location: at(7, "…theo Khoản 2 Điều 5…"), fix: { kind: "edits", index: 7, edits: [{ find: "Khoản 2", replacement: "khoản 2" }], label: "Viết thường: khoản 2" } },
  { ruleId: "footer.hoTen.bold", severity: "error", message: "Họ và tên người ký phải in đậm.", reference: "Phụ lục I", location: at(15, "Nguyễn Văn A"), fix: { kind: "font", index: 15, bold: true, label: "In đậm họ tên" } },
];

mountApp({ check: () => Promise.resolve(SAMPLE_ISSUES) });
setStatus("Xem trước giao diện (không có Word).");
