import { buildFooterBlock, type FooterBlockInput } from "./footerBlock";
import { buildHeaderBlock, type HeaderBlockInput } from "./headerBlock";
import { buildTitleBody, type TitleBlockInput } from "./titleBlock";

/** Chữ đánh dấu chỗ nội dung; lớp Word thay bằng đoạn trống đúng style rồi đặt con trỏ vào đó. */
export const BODY_MARKER = "§NOI-DUNG§";

/** Đoạn chỗ nội dung: có thể gắn sẵn style và giãn dòng Exactly để người dùng gõ tiếp đúng thể thức. */
export interface BodyParagraph {
  /** Mã style (styleId) của style nội dung đã cài; Word bỏ dấu gạch dưới khỏi tên style khi tạo mã. */
  styleId?: string;
  /** Giãn dòng Exactly (pt) đặt trực tiếp trên đoạn (chuẩn Đảng); style cài bằng JSON chỉ ra "tối thiểu". */
  exactPt?: number;
}

export function markerParagraph(body: BodyParagraph = {}): string {
  const style = body.styleId ? `<w:pStyle w:val="${body.styleId}"/>` : "";
  const spacing = body.exactPt === undefined ? "" : `<w:spacing w:line="${Math.round(body.exactPt * 20)}" w:lineRule="exact"/>`;
  const props = style + spacing === "" ? "" : `<w:pPr>${style}${spacing}</w:pPr>`;
  return `<w:p>${props}<w:r><w:t>${BODY_MARKER}</w:t></w:r></w:p>`;
}

export interface AssembleInput {
  header: HeaderBlockInput;
  /** Tên loại, trích yếu hoặc Kính gửi; bỏ trống thì không có phần này. */
  title?: TitleBlockInput;
  footer: FooterBlockInput;
}

/**
 * Thân văn bản của một công văn hoàn chỉnh: khối đầu, tên loại/trích yếu/Kính gửi, chỗ nội dung, khối cuối.
 * Mỗi khối do bộ dựng riêng của nó chịu trách nhiệm (T7 đến T10); ở đây chỉ ghép theo thứ tự.
 */
export function assembleDocumentBody(input: AssembleInput, body: BodyParagraph = {}): string {
  const middle = input.title ? buildTitleBody(input.title) : "<w:p/>";
  return buildHeaderBlock(input.header) + middle + markerParagraph(body) + buildFooterBlock(input.footer);
}
