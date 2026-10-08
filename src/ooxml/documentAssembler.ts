import { buildFooterBlock, type FooterBlockInput } from "./footerBlock";
import { buildHeaderBlock, type HeaderBlockInput } from "./headerBlock";
import { buildTitleBody, type TitleBlockInput } from "./titleBlock";

/** Chữ đánh dấu chỗ nội dung; lớp Word thay bằng đoạn trống đúng style rồi đặt con trỏ vào đó. */
export const BODY_MARKER = "§NOI-DUNG§";

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
export function assembleDocumentBody(input: AssembleInput): string {
  const middle = input.title ? buildTitleBody(input.title) : "<w:p/>";
  const marker = `<w:p><w:r><w:t>${BODY_MARKER}</w:t></w:r></w:p>`;
  return buildHeaderBlock(input.header) + middle + marker + buildFooterBlock(input.footer);
}
