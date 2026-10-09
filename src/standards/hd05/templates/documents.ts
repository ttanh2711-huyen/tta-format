import { markerParagraph, type BodyParagraph } from "../../../ooxml/documentAssembler";
import { escapeXml } from "../../../ooxml/headerBlock";
import { buildMultiSignBlock, type MultiSignInput } from "../../../ooxml/multiSignBlock";
import { buildMinutesBody } from "../../../templates/bienBan";
import { buildSmallFormBody, type SmallForm } from "../../../templates/smallForms";
import { buildPartyFooterBlock } from "../footerBlock";
import { buildPartyHeaderBlock } from "../headerBlock";
import type { PartyAssembleInput } from "../assembleDocument";
import { buildPartyTitleBody } from "../titleBlock";

/**
 * Các họ mẫu còn lại của chuẩn Đảng (TD7): biên bản (hai chữ ký, có dòng xác nhận chữ ký khi không đóng dấu), văn bản đại
 * hội (tên đại hội trên thể loại, ký hiệu "ĐH"), văn bản liên cơ quan (hai khối ký) và giấy tờ hành chính nhỏ. Mỗi họ
 * ghép khối đầu, thể loại và trích yếu của TD1 đến TD3 với phần thân và khối cuối riêng.
 */

type Head = Pick<PartyAssembleInput, "header" | "title">;

function titled(input: Head): { head: string; title: string } {
  if (!input.title || input.title.variant === "congVan") throw new Error("Văn bản này phải có thể loại và trích yếu.");
  return { head: buildPartyHeaderBlock(input.header), title: buildPartyTitleBody(input.title) };
}

/** Biên bản: khối đầu, thể loại, các mục (Thời gian, Địa điểm...), khối ký nhiều bên (người ghi biên bản và chủ trì hội nghị). */
export function assemblePartyMinutesBody(input: Head, stylePrefix: string, sign: MultiSignInput): string {
  const { head, title } = titled(input);
  return head + title + buildMinutesBody(stylePrefix) + buildMultiSignBlock(sign);
}

/** Văn bản liên cơ quan: hai khối ký cạnh nhau (tên cơ quan nối bằng gạch nối ở khối đầu và ký hiệu). */
export function assemblePartyJointBody(input: Head, sign: MultiSignInput, body: BodyParagraph = {}): string {
  const { head, title } = titled(input);
  if (sign.parties.length !== 2) throw new Error("Văn bản liên cơ quan có đúng hai khối ký.");
  return head + title + markerParagraph(body) + buildMultiSignBlock(sign);
}

/** Dòng tên đại hội đặt trên thể loại: in hoa đậm, canh giữa. */
export function buildCongressName(name: string, sizePt: number): string {
  const text = name.replace(/\s+/g, " ").trim().toUpperCase();
  if (text === "") throw new Error("Hãy nhập tên đại hội.");
  const half = sizePt * 2;
  return `<w:p><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

/** Văn bản đại hội: khối đầu, tên đại hội, thể loại và trích yếu, chỗ nội dung, khối cuối thường (có thể kèm dòng xác nhận chữ ký). */
export function assemblePartyCongressBody(input: PartyAssembleInput, congressName: string, sizePt: number, body: BodyParagraph = {}): string {
  const { head, title } = titled(input);
  return head + buildCongressName(congressName, sizePt) + title + markerParagraph(body) + buildPartyFooterBlock(input.footer);
}

/** Giấy tờ hành chính nhỏ: khối đầu, thể loại và trích yếu, các mục cần điền, khối cuối. */
export function assemblePartyFormBody(input: PartyAssembleInput, form: SmallForm, stylePrefix: string): string {
  const { head, title } = titled(input);
  return head + title + buildSmallFormBody(form, stylePrefix) + buildPartyFooterBlock(input.footer);
}

