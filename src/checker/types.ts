import type { ParagraphModel } from "../rules/docModel";

export type Severity = "error" | "warning";

/** Nơi lỗi nằm: một phần (trang) của tài liệu hoặc một đoạn. */
export type IssueLocation =
  | { kind: "page"; section: number }
  | { kind: "paragraph"; index: number; excerpt: string };

/**
 * Cách sửa một lỗi: thay chữ trong một đoạn (giữ định dạng của chữ) hoặc đặt lại định dạng chữ của cả đoạn.
 * `index` là chỉ số đoạn trong `body.paragraphs`; chỉ đụng đoạn đó.
 */
export type Fix =
  | { kind: "replace"; index: number; find: string; replacement: string; label: string }
  | { kind: "font"; index: number; bold?: boolean; italic?: boolean; sizePt?: number; label: string };

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
