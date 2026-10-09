import type { AssembleInput } from "../ooxml/documentAssembler";
import { BODY_MARKER } from "../ooxml/documentAssembler";
import { buildHeaderBlock, escapeXml } from "../ooxml/headerBlock";
import { buildMultiSignBlock, type MultiSignInput } from "../ooxml/multiSignBlock";
import { styleIdFromName } from "../ooxml/styleXml";
import { buildTitleBody } from "../ooxml/titleBlock";

/**
 * Biên bản (T19, Mẫu 1.9): tên loại "BIÊN BẢN" và trích yếu, các mục Thời gian, Địa điểm, Thành phần, Chủ trì, Thư ký,
 * Nội dung, rồi khối ký nhiều bên (mặc định chủ tọa và thư ký), không bắt buộc "Nơi nhận". Con trỏ nằm sau nhãn đầu tiên.
 */

/** Nhãn các mục đầu biên bản, theo thứ tự. */
export const MINUTES_FIELDS: readonly string[] = ["Thời gian:", "Địa điểm:", "Thành phần tham dự:", "Chủ trì (chủ tọa):", "Thư ký:", "Nội dung:"];

function field(label: string, styleId: string, marker: boolean): string {
  const run = marker ? `<w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>${BODY_MARKER}</w:t></w:r>` : "";
  return `<w:p><w:pPr><w:pStyle w:val="${styleId}"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(label)}</w:t></w:r>${run}</w:p>`;
}

/** Các mục đầu biên bản bằng style nội dung của chuẩn; chữ đánh dấu chỗ nội dung nằm sau nhãn đầu. */
export function buildMinutesBody(stylePrefix: string): string {
  const style = styleIdFromName(`${stylePrefix}NoiDung`);
  return MINUTES_FIELDS.map((label, index) => field(label, style, index === 0)).join("");
}

/** Hai bên ký mặc định của biên bản. */
export function defaultMinutesParties(): MultiSignInput["parties"] {
  return [
    { title: "Chủ tọa", positionLines: [], signerName: "" },
    { title: "Thư ký", positionLines: [], signerName: "" },
  ];
}

/** Thân văn bản đầy đủ: khối đầu, tên loại, các mục biên bản, khối ký nhiều bên. */
export function assembleMinutesBody(input: Pick<AssembleInput, "header" | "title">, stylePrefix: string, sign: MultiSignInput): string {
  if (!input.title || input.title.variant !== "named") throw new Error("Biên bản phải dùng khối tên loại và trích yếu.");
  return buildHeaderBlock(input.header) + buildTitleBody(input.title) + buildMinutesBody(stylePrefix) + buildMultiSignBlock(sign);
}
