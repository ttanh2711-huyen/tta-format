import { markerParagraph, type BodyParagraph } from "../../ooxml/documentAssembler";
import { buildPartyFooterBlock, type PartyFooterInput } from "./footerBlock";
import { buildPartyHeaderBlock, type PartyHeaderInput } from "./headerBlock";
import { buildPartyTitleBody, type PartyTitleInput } from "./titleBlock";

export interface PartyAssembleInput {
  header: PartyHeaderInput;
  /** Thể loại, trích yếu, Kính gửi/Kính trình, căn cứ; bỏ trống thì không có phần này. */
  title?: PartyTitleInput;
  footer: PartyFooterInput;
}

/**
 * Thân văn bản Đảng hoàn chỉnh (TD5): khối đầu, thể loại và trích yếu hoặc Kính gửi, căn cứ, chỗ nội dung, khối
 * cuối. Mỗi khối do bộ dựng riêng của nó (TD1 đến TD4) chịu trách nhiệm; ở đây chỉ ghép theo thứ tự.
 */
export function assemblePartyDocumentBody(input: PartyAssembleInput, body: BodyParagraph = {}): string {
  const middle = input.title ? buildPartyTitleBody(input.title) : "<w:p/>";
  return buildPartyHeaderBlock(input.header) + middle + markerParagraph(body) + buildPartyFooterBlock(input.footer);
}
