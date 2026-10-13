/**
 * Giả định các mục còn thiếu (T26): người dùng chưa điền khối đầu hoặc khối cuối vẫn tạo được văn bản đúng mẫu của chuẩn;
 * các mục bắt buộc còn trống được điền chữ giả định để người dùng sửa ngay trong văn bản. Hàm ở đây thuần tuý (không chạm DOM)
 * nên kiểm thử được; biểu mẫu gọi chúng khi bấm tạo văn bản.
 */

export const PLACEHOLDERS = {
  issuer: "TÊN CƠ QUAN BAN HÀNH",
  place: "Địa danh",
  agencyAbbr: "CQ",
  subject: "nội dung văn bản",
  recipient: "Cơ quan, tổ chức nhận văn bản",
  typeName: "Tên loại văn bản",
  position: "Chức vụ",
  signer: "Họ và tên",
  authority: "Thẩm quyền ban hành",
  basis: "Căn cứ pháp lý ban hành",
  congress: "Đại hội đại biểu lần thứ …",
} as const;

export interface HeaderDraft {
  issuer: string;
  place: string;
  agencyAbbrs: string;
  subject: string;
  recipients: string;
  typeName: string;
  /** Văn bản có "Kính gửi" hoặc "Kính trình" cần nơi nhận. */
  needsRecipients: boolean;
  /** Văn bản có tên loại (không phải công văn) cần tên loại. */
  needsTypeName: boolean;
  withNumberRow: boolean;
  withTitle: boolean;
}

export interface FooterDraft {
  recipients: string;
  positionLines: readonly string[];
  signerName: string;
}

export interface Assumed<T> {
  draft: T;
  /** Tên các mục đã điền chữ giả định, theo thứ tự. */
  assumed: string[];
}

const blank = (text: string): boolean => text.replace(/\s+/g, "") === "";

/** Điền chữ giả định vào các mục bắt buộc còn trống của khối đầu; không bật `assume` thì giữ nguyên. */
export function assumeHeader(draft: HeaderDraft, assume: boolean): Assumed<HeaderDraft> {
  if (!assume) return { draft, assumed: [] };
  const assumed: string[] = [];
  const next = { ...draft };
  const fill = (key: "issuer" | "place" | "agencyAbbrs" | "subject" | "recipients" | "typeName", value: string, label: string): void => {
    if (blank(next[key])) {
      next[key] = value;
      assumed.push(label);
    }
  };
  fill("issuer", PLACEHOLDERS.issuer, "cơ quan ban hành");
  if (draft.withNumberRow) {
    fill("place", PLACEHOLDERS.place, "địa danh");
    fill("agencyAbbrs", PLACEHOLDERS.agencyAbbr, "chữ viết tắt cơ quan");
  }
  if (draft.withTitle) {
    if (draft.needsTypeName) fill("typeName", PLACEHOLDERS.typeName, "tên loại văn bản");
    fill("subject", PLACEHOLDERS.subject, "trích yếu");
    if (draft.needsRecipients) fill("recipients", PLACEHOLDERS.recipient, "nơi nhận Kính gửi");
  }
  return { draft: next, assumed };
}

/** Điền chữ giả định vào các mục bắt buộc còn trống của khối cuối (chức vụ, họ tên người ký). */
export function assumeFooter(draft: FooterDraft, assume: boolean): Assumed<FooterDraft> {
  if (!assume) return { draft, assumed: [] };
  const assumed: string[] = [];
  let positionLines = draft.positionLines;
  let signerName = draft.signerName;
  if (positionLines.every(blank)) {
    positionLines = [PLACEHOLDERS.position];
    assumed.push("chức vụ người ký");
  }
  if (blank(signerName)) {
    signerName = PLACEHOLDERS.signer;
    assumed.push("họ tên người ký");
  }
  return { draft: { ...draft, positionLines, signerName }, assumed };
}
