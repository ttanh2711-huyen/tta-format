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

/**
 * Các mục riêng của từng loại biên bản (theo mã loại). Biên bản sinh hoạt chi bộ (T27): bố cục theo cách ghi thông dụng
 * trong hướng dẫn sinh hoạt chi bộ; là đề xuất, cần đối chiếu hướng dẫn hiện hành của cấp ủy cấp trên.
 */
export const MINUTES_VARIANTS: Readonly<Record<string, readonly string[]>> = {
  "bien-ban-sinh-hoat-chi-bo": [
    "Thời gian: bắt đầu lúc … giờ … phút, ngày … tháng … năm …",
    "Địa điểm:",
    "Chủ trì:",
    "Thư ký:",
    "Đảng viên của chi bộ: tổng số …; có mặt …; vắng mặt … (có lý do …, không lý do …).",
    "Đại biểu dự (nếu có):",
    "I. Đánh giá kết quả thực hiện nghị quyết tháng trước:",
    "II. Nhiệm vụ, giải pháp tháng tới:",
    "III. Ý kiến thảo luận của đảng viên:",
    "IV. Kết luận của đồng chí chủ trì:",
    "Biểu quyết thông qua nghị quyết: tán thành …/…, không tán thành …/…",
    "Buổi sinh hoạt kết thúc lúc … giờ … phút cùng ngày.",
  ],
};

/** Các bên ký mặc định theo loại biên bản (nếu khác mặc định). */
export const MINUTES_SIGNERS: Readonly<Record<string, readonly string[]>> = {
  "bien-ban-sinh-hoat-chi-bo": ["Thư ký", "Chủ trì"],
};

function field(label: string, styleId: string, marker: boolean): string {
  const run = marker ? `<w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>${BODY_MARKER}</w:t></w:r>` : "";
  return `<w:p><w:pPr><w:pStyle w:val="${styleId}"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(label)}</w:t></w:r>${run}</w:p>`;
}

/** Các mục đầu biên bản bằng style nội dung của chuẩn; chữ đánh dấu chỗ nội dung nằm sau nhãn đầu. */
export function buildMinutesBody(stylePrefix: string, fields: readonly string[] = MINUTES_FIELDS): string {
  const style = styleIdFromName(`${stylePrefix}NoiDung`);
  return fields.map((label, index) => field(label, style, index === 0)).join("");
}

/** Hai bên ký mặc định của biên bản. */
export function defaultMinutesParties(): MultiSignInput["parties"] {
  return [
    { title: "Chủ tọa", positionLines: [], signerName: "" },
    { title: "Thư ký", positionLines: [], signerName: "" },
  ];
}

/** Thân văn bản đầy đủ: khối đầu, tên loại, các mục biên bản, khối ký nhiều bên. */
export function assembleMinutesBody(input: Pick<AssembleInput, "header" | "title">, stylePrefix: string, sign: MultiSignInput, fields?: readonly string[]): string {
  if (!input.title || input.title.variant !== "named") throw new Error("Biên bản phải dùng khối tên loại và trích yếu.");
  return buildHeaderBlock(input.header) + buildTitleBody(input.title) + buildMinutesBody(stylePrefix, fields) + buildMultiSignBlock(sign);
}
