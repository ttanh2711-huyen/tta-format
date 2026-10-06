/**
 * Kiểu dữ liệu dùng chung cho các chuẩn văn bản. Mỗi chuẩn (NĐ 30, HD 05) khai báo
 * một `StandardProfile`; mã dùng chung (kiểm tra, định dạng, giao diện) chỉ nhận
 * `StandardProfile`, không rẽ nhánh theo mã chuẩn.
 */

export interface Range {
  min: number;
  max: number;
}

export type Casing = "upper" | "lower"; // chữ in hoa | chữ in thường
export type FontStyle = "upright" | "italic"; // kiểu chữ đứng | nghiêng

export type MarginSide = "top" | "bottom" | "left" | "right";

export interface ElementRule<Id extends string = string> {
  id: Id;
  label: string;
  casing: Casing;
  sizePt: Range;
  style: FontStyle;
  bold: boolean;
  /** Gạch chân (chuẩn Đảng: cụm từ "Nơi nhận"). */
  underline?: boolean;
  /** Vị trí trích dẫn trong văn bản gốc. */
  source: string;
}

export interface MarginRange extends Range {
  source: string;
}

export interface PageRules {
  paper: { name: "A4"; widthMm: number; heightMm: number; source: string };
  /** Lề cách mép giấy, mm. Chuẩn cố định một giá trị thì min bằng max. */
  marginsMm: Record<MarginSide, MarginRange>;
  /** Lề mặt sau khi in hai mặt; không có nghĩa là chuẩn không quy định riêng. */
  duplexBackMarginsMm?: Record<MarginSide, MarginRange>;
  /** Cho phép trình bày theo chiều ngang khi nội dung có bảng, biểu. */
  landscapeAllowedForTables: boolean;
}

export interface FontRules {
  name: string;
  encoding: string;
  color: string;
  source: string;
}

export interface PageNumberRules {
  sizePt: Range;
  align: "center";
  position: "top-margin";
  /** Khoảng cách từ mép trên giấy đến số trang, nếu chuẩn quy định. */
  distanceFromTopEdgeMm?: number;
  hideOnFirstPage: boolean;
  appendixNumberedSeparately: boolean;
  source: string;
}

export type LineSpacingRule =
  | { kind: "multiple"; min: number; max: number } // bội số dòng đơn (lines)
  | { kind: "exactly"; minPt: number; maxPt: number }; // khoảng cách cố định (Exactly)

export interface ContentLayout {
  align: "justify";
  /** Các giá trị thụt dòng đầu hợp lệ, mm. */
  firstLineIndentMm: readonly number[];
  /** Sai số cho phép khi so thụt dòng đầu, mm. */
  firstLineIndentToleranceMm: number;
  paragraphSpacingMinPt: number;
  lineSpacing: LineSpacingRule;
  /** Nội dung văn bản kết thúc bằng dấu chấm (chỉ chuẩn nào quy định mới có). */
  endsWithPeriod?: true;
  source: string;
}

export interface SignAuthority {
  abbr: string;
  meaning: string;
}

export interface UrgencyMark {
  text: string;
  /** Kích thước con dấu (mm), nếu chuẩn quy định. */
  widthMm?: number;
  heightMm?: number;
}

export interface DocType<Layout extends string = string> {
  id: string;
  name: string;
  /** Chữ viết tắt tên loại dùng trong ký hiệu văn bản; null nếu không có. */
  abbr: string | null;
  layout: Layout;
  /** Số mẫu trong văn bản gốc; null nếu không có mẫu. */
  template: string | null;
  /** Nguồn của chữ viết tắt: văn bản gốc nêu rõ hay suy ra theo quy tắc chữ cái đầu. */
  abbrSource?: "explicit" | "derived";
  note?: string;
}

/** Giá trị trang mặc định khi bấm "Thiết lập trang"; mọi giá trị phải hợp lệ với chính chuẩn. */
export interface PageDefaults {
  marginsMm: Record<MarginSide, number>;
  headerDistanceMm: number;
  footerDistanceMm: number;
  source: string;
}

export type StandardId = "nd30" | "hd05";

export interface StandardProfile {
  id: StandardId;
  /** Tên hiển thị trên giao diện. */
  name: string;
  legalBasis: string;
  /** Ngày có hiệu lực của văn bản gốc, dạng ISO (YYYY-MM-DD). */
  effectiveDate: string;
  /** Dòng chữ nhận diện chuẩn ở đầu văn bản (dùng khi tự nhận diện chuẩn của tài liệu). */
  identifyingTitle: string;
  page: PageRules;
  pageDefaults: PageDefaults;
  font: FontRules;
  pageNumber: PageNumberRules;
  content: ContentLayout;
  elements: Readonly<Record<string, ElementRule>>;
  docTypes: readonly DocType[];
  signAuthority: readonly SignAuthority[];
  urgencyMarks: readonly UrgencyMark[];
  /**
   * Dòng số và ký hiệu văn bản theo cách ghi của chuẩn.
   * `typeAbbr`: chữ viết tắt tên loại (null với công văn Chính quyền); `agencyAbbrs`: chữ viết tắt
   * tên cơ quan, đơn vị (nhiều phần tử với văn bản liên cơ quan hoặc có đơn vị soạn thảo).
   */
  formatSoKyHieu(so: number, typeAbbr: string | null, agencyAbbrs: readonly string[]): string;
}
