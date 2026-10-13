import type { ParagraphModel } from "../rules/docModel";

export type Severity = "error" | "warning";

/** Nơi lỗi nằm: một phần (trang) của tài liệu hoặc một đoạn. */
export type IssueLocation =
  | { kind: "page"; section: number }
  | { kind: "paragraph"; index: number; excerpt: string };

/**
 * Cách sửa một lỗi. Các cách sửa trên đoạn (`index` là chỉ số đoạn trong `body.paragraphs`) chỉ đụng đoạn đó:
 * thay chữ giữ định dạng của chữ, đặt phông, đậm, nghiêng, cỡ chữ, căn lề, thụt, giãn đoạn, giãn dòng bội số, hoặc
 * giãn dòng Exactly (qua OOXML). Cách sửa `page` đặt khổ giấy, lề, khoảng cách header của một phần.
 */
export type Fix =
  | { kind: "replace"; index: number; find: string; replacement: string; label: string }
  | { kind: "font"; index: number; name?: string; bold?: boolean; italic?: boolean; underline?: boolean; sizePt?: number; label: string }
  | { kind: "paragraph"; index: number; align?: "both"; firstLineIndentMm?: number; spaceBeforePt?: number; spaceAfterPt?: number; lineMultiple?: number; label: string }
  | { kind: "exact"; index: number; pt: number; label: string }
  | { kind: "insertAfter"; index: number; text: string; align?: "center" | "left"; sizePt?: number; bold?: boolean; italic?: boolean; label: string }
  | { kind: "page"; section: number; paper?: "A4"; marginsMm?: Partial<Record<"top" | "bottom" | "left" | "right", number>>; headerDistanceMm?: number; label: string };

export interface Issue {
  /** Mã quy tắc ổn định, ví dụ "page.margin.left". */
  ruleId: string;
  severity: Severity;
  /** Câu mô tả lỗi bằng tiếng Việt. */
  message: string;
  /** Điều khoản tham chiếu của chuẩn. */
  reference: string;
  actual?: string;
  expected?: string;
  location: IssueLocation;
  /** Có thì lỗi sửa được tự động bằng nút "Sửa". */
  fix?: Fix;
}

/** Thiết lập trang của một phần, đọc từ Word. */
export interface PageInfo {
  widthMm: number;
  heightMm: number;
  marginsMm: { top: number; bottom: number; left: number; right: number };
  headerDistanceMm: number;
  footerDistanceMm: number;
  mirrorMargins: boolean;
}

export interface CheckParagraph {
  /** Thứ tự trong body.paragraphs của Word (kể cả đoạn trống và đoạn trong bảng). */
  index: number;
  inTable: boolean;
  /** Vị trí trong bảng: bảng thứ mấy, hàng, cột (đếm từ 0); không có nếu đoạn ngoài bảng. */
  table?: { index: number; row: number; col: number };
  /** Tên style hiển thị, nếu biết. */
  styleName?: string;
  /** Đoạn với định dạng hiệu lực đã giải style. */
  resolved: ParagraphModel;
}

export interface CheckDocument {
  pages: PageInfo[];
  paragraphs: CheckParagraph[];
}
