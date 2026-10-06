import { STANDARDS } from "../standards";
import type { StandardProfile } from "../standards/types";

/** Thông tin một đoạn đọc từ Word, đủ để quyết định có áp style nội dung hay không. */
export interface ParagraphInfo {
  text: string;
  /** Tên style hiển thị (theo ngôn ngữ giao diện của Word). */
  styleName: string;
  /** Tên style có sẵn không phụ thuộc ngôn ngữ, ví dụ "Heading1", "Normal"; "Other" nếu tự đặt. */
  styleBuiltIn: string;
  tableNestingLevel: number;
  isListItem: boolean;
}

export type SkipReason = "empty" | "table" | "standardStyle" | "list" | "heading";

export type Classification = { apply: true } | { apply: false; reason: SkipReason };

const BUILT_IN_HEADINGS = /^(Heading[1-9]|Title|Subtitle)$/;

/**
 * Quyết định một đoạn có được áp style nội dung hay không. Bỏ qua: đoạn trong bảng, đoạn trống, đoạn đã
 * dùng style của bất kỳ bộ chuẩn nào, tiêu đề có sẵn của Word và mục danh sách tự động (áp style có thể
 * làm hỏng đánh số).
 */
export function classifyParagraph(info: ParagraphInfo): Classification {
  if (info.tableNestingLevel > 0) return { apply: false, reason: "table" };
  if (info.text.trim() === "") return { apply: false, reason: "empty" };
  if (STANDARDS.some((standard) => info.styleName.startsWith(standard.styles.prefix))) {
    return { apply: false, reason: "standardStyle" };
  }
  if (BUILT_IN_HEADINGS.test(info.styleBuiltIn)) return { apply: false, reason: "heading" };
  if (info.isListItem) return { apply: false, reason: "list" };
  return { apply: true };
}

export type SkippedCounts = Record<SkipReason, number>;

export interface FormatPlan {
  /** Chỉ số các đoạn sẽ được áp style. */
  targets: number[];
  skipped: SkippedCounts;
  /** Tên style nội dung sẽ áp, ví dụ "ND30_NoiDung". */
  styleName: string;
}

export function planFormatBody(paragraphs: readonly ParagraphInfo[], profile: StandardProfile): FormatPlan {
  const skipped: SkippedCounts = { empty: 0, table: 0, standardStyle: 0, list: 0, heading: 0 };
  const targets: number[] = [];
  paragraphs.forEach((info, index) => {
    const result = classifyParagraph(info);
    if (result.apply) {
      targets.push(index);
    } else {
      skipped[result.reason] += 1;
    }
  });
  return { targets, skipped, styleName: `${profile.styles.prefix}NoiDung` };
}

const SKIP_LABEL: ReadonlyArray<[SkipReason, string]> = [
  ["empty", "đoạn trống"],
  ["table", "đoạn trong bảng"],
  ["standardStyle", "đoạn đã dùng style của bộ chuẩn"],
  ["list", "mục danh sách tự động"],
  ["heading", "tiêu đề có sẵn của Word"],
];

export interface FormatSummaryInput {
  applied: number;
  skipped: SkippedCounts;
  notes: readonly string[];
  /** Có ghi chú cần cảnh báo (ví dụ chưa đặt được Exactly). */
  warn?: boolean;
}

/** "Bỏ qua 6 đoạn: 2 đoạn trống, 3 đoạn trong bảng."; chuỗi rỗng nếu không bỏ qua đoạn nào. */
export function describeSkipped(skipped: SkippedCounts): string {
  const parts = SKIP_LABEL.filter(([reason]) => skipped[reason] > 0).map(([reason, label]) => `${skipped[reason]} ${label}`);
  const total = SKIP_LABEL.reduce((sum, [reason]) => sum + skipped[reason], 0);
  return total > 0 ? `Bỏ qua ${total} đoạn: ${parts.join(", ")}.` : "";
}

export function summarizeFormat(input: FormatSummaryInput): { kind: "ok" | "warning"; text: string } {
  const { applied, skipped, notes } = input;
  const head = applied > 0 ? `Đã định dạng ${applied} đoạn.` : "Không có đoạn nào để định dạng.";
  const skipText = describeSkipped(skipped);
  const text = [head, skipText, ...notes].filter((part) => part !== "").join(" ");
  return { kind: applied === 0 || input.warn ? "warning" : "ok", text };
}

export interface TableInfo {
  /** 1 là bảng ngoài cùng, 2 trở lên là bảng lồng nhau. */
  nestingLevel: number;
  /** Kiểu viền ngoài và viền trong của bảng (trên, dưới, trái, phải, ngang trong, dọc trong). */
  borderTypes: readonly string[];
}

export type TableKind = "data" | "layout" | "nested";

/**
 * Phân loại bảng để chỉ áp style chữ bảng cho bảng dữ liệu. Bảng không viền thường là bảng bố cục
 * (đầu văn bản, nơi nhận và chữ ký) nên không đụng tới; không đọc được viền thì không đoán.
 */
export function classifyTable(info: TableInfo): TableKind {
  if (info.nestingLevel > 1) return "nested";
  if (info.borderTypes.length === 0) return "layout";
  return info.borderTypes.every((type) => type === "None") ? "layout" : "data";
}
