import { bodyPart, hasPageField, readJustification } from "./facts";

/**
 * Đánh số trang (T11): trường PAGE canh giữa ở header mặc định, trang đầu khác biệt để trang 1 không hiện số.
 * Phần thuần ở đây dựng OOXML và lập kế hoạch từng phần; phần gọi Word ở `office/pageNumber.ts`.
 */

const FONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>';

/** Đoạn canh giữa chứa trường số trang PAGE, cỡ chữ `sizePt`. */
export function pageNumberParagraph(sizePt: number): string {
  const half = sizePt * 2;
  return (
    `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="center"/></w:pPr>` +
    `<w:fldSimple w:instr=" PAGE "><w:r><w:rPr>${FONT}<w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p>`
  );
}

export interface PageNumberState {
  hasPage: boolean;
  centered: boolean;
  sizePt: number | undefined;
}

/**
 * Đọc trạng thái số trang từ OOXML của một header. Chỉ đọc phần thân của gói (nội dung header); gói Word trả
 * về còn kèm phần styles, nếu quét cả gói sẽ đọc nhầm cỡ chữ và căn lề của style.
 */
export function readPageNumberState(headerXml: string): PageNumberState {
  const xml = bodyPart(headerXml) || headerXml;
  // Đánh giá đúng đoạn chứa trường PAGE, để đoạn trống đứng trước hoặc sau không làm sai cỡ chữ và căn lề.
  const paragraphs = xml.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g) ?? [];
  const holder = paragraphs.find((paragraph) => hasPageField(paragraph));
  const scope = holder ?? xml;
  const size = /<w:sz\s+w:val="(\d+)"/.exec(scope)?.[1];
  return {
    hasPage: holder !== undefined || hasPageField(xml),
    centered: readJustification(scope) === "center",
    sizePt: size === undefined ? undefined : Number(size) / 2,
  };
}

/** Mô tả ngắn trạng thái đọc được, dùng trong thông báo khi đọc lại chưa đúng. */
export function describePageNumberState(state: PageNumberState): string {
  return `có trường PAGE: ${state.hasPage ? "có" : "không"}, căn giữa: ${state.centered ? "có" : "không"}, cỡ: ${state.sizePt ?? "không rõ"}`;
}

export interface SectionState {
  differentFirst: boolean;
  primary: PageNumberState;
  /** Header trang đầu: có trường PAGE không, và có trống không. */
  first: { hasPage: boolean; empty: boolean };
}

export interface PageNumberOptions {
  sizePt: number;
  /** Trang 1 không hiện số. */
  hideFirst: boolean;
}

export interface SectionPlan {
  /** Đặt lại "trang đầu khác biệt"; undefined nếu đã đúng. */
  setDifferentFirst: boolean | undefined;
  /** Ghi lại số trang ở header mặc định. */
  writePrimary: boolean;
  /** Xoá số trang đang nằm ở header trang đầu. */
  clearFirst: boolean;
}

/**
 * Việc cần làm cho từng phần. Chỉ phần đầu tiên ẩn số ở trang đầu; các phần sau luôn hiện số từ trang
 * đầu của phần. Mỗi phần chỉ bị ghi khi chưa đúng, nên chạy lại không đổi gì và không tạo trùng trường số trang.
 */
export function planPageNumbers(sections: readonly SectionState[], options: PageNumberOptions): SectionPlan[] {
  return sections.map((section, index) => {
    const wantDifferent = index === 0 && options.hideFirst;
    const primary = section.primary;
    const primaryOk = primary.hasPage && primary.centered && primary.sizePt === options.sizePt;
    return {
      setDifferentFirst: section.differentFirst === wantDifferent ? undefined : wantDifferent,
      writePrimary: !primaryOk,
      clearFirst: wantDifferent && section.first.hasPage,
    };
  });
}

export function isPlanEmpty(plan: readonly SectionPlan[]): boolean {
  return plan.every((item) => item.setDifferentFirst === undefined && !item.writePrimary && !item.clearFirst);
}
