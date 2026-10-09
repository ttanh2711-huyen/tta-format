import type { AssembleInput } from "../ooxml/documentAssembler";
import { BODY_MARKER } from "../ooxml/documentAssembler";
import { buildFooterBlock } from "../ooxml/footerBlock";
import { buildHeaderBlock, escapeXml } from "../ooxml/headerBlock";
import { styleIdFromName } from "../ooxml/styleXml";
import { buildTitleBody } from "../ooxml/titleBlock";

/**
 * Giấy tờ và mẫu nhỏ (T20): giấy mời (Mẫu 1.7), giấy giới thiệu (Mẫu 1.8), công điện (Mẫu 1.6), giấy nghỉ phép
 * (Mẫu 1.10), thư công (không có mẫu). Mỗi mẫu là khối đầu, tên loại và trích yếu, các mục cần điền (nhãn đậm) và khối
 * cuối thường. LƯU Ý: các mục và câu kết là bố cục theo cách trình bày thông dụng, chưa đối chiếu nguyên văn Phụ lục III;
 * giao diện ghi rõ điều này cho từng mẫu.
 */

export interface SmallForm {
  /** Mã loại trong `docTypes`. */
  docTypeId: string;
  /** Các mục cần điền, nhãn in đậm, theo thứ tự; con trỏ đặt ở mục đầu. */
  fields: readonly string[];
  /** Câu kết (chữ thường), trống thì bỏ. */
  closing: string;
  /** Nguồn bố cục, hiện cho người dùng. */
  source: string;
}

const UNVERIFIED = "bố cục theo cách trình bày thông dụng, chưa đối chiếu nguyên văn";

export const SMALL_FORMS: readonly SmallForm[] = [
  { docTypeId: "giay-moi", fields: ["Kính mời:", "Đến dự:", "Thời gian:", "Địa điểm:", "Nội dung:"], closing: "Rất mong sự có mặt của quý vị.", source: `Mẫu 1.7, ${UNVERIFIED} Phụ lục III` },
  { docTypeId: "giay-gioi-thieu", fields: ["Kính gửi:", "Trân trọng giới thiệu ông (bà):", "Chức vụ:", "Được cử đến:", "Về việc:"], closing: "Đề nghị quý cơ quan tạo điều kiện để ông (bà) có tên trên hoàn thành nhiệm vụ.", source: `Mẫu 1.8, ${UNVERIFIED} Phụ lục III` },
  { docTypeId: "cong-dien", fields: ["Kính gửi:", "Nội dung:"], closing: "", source: `Mẫu 1.6, ${UNVERIFIED} Phụ lục III` },
  { docTypeId: "giay-nghi-phep", fields: ["Cấp cho ông (bà):", "Chức vụ:", "Đơn vị công tác:", "Được nghỉ phép từ ngày … đến ngày …:", "Nơi nghỉ phép:"], closing: "", source: `Mẫu 1.10, ${UNVERIFIED} Phụ lục III` },
  { docTypeId: "thu-cong", fields: ["Kính gửi:", "Nội dung:"], closing: "Trân trọng.", source: "Phụ lục III không có mẫu: bố cục đề xuất, cần bạn duyệt" },
];

export function findSmallForm(docTypeId: string): SmallForm | undefined {
  return SMALL_FORMS.find((form) => form.docTypeId === docTypeId);
}

function field(label: string, styleId: string, marker: boolean): string {
  const mark = marker ? `<w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>${BODY_MARKER}</w:t></w:r>` : "";
  return `<w:p><w:pPr><w:pStyle w:val="${styleId}"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${escapeXml(label)}</w:t></w:r>${mark}</w:p>`;
}

/** XML các mục của mẫu bằng style nội dung; chữ đánh dấu chỗ nội dung nằm sau nhãn đầu. */
export function buildSmallFormBody(form: SmallForm, stylePrefix: string): string {
  const style = styleIdFromName(`${stylePrefix}NoiDung`);
  const closing = form.closing === "" ? "" : `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(form.closing)}</w:t></w:r></w:p>`;
  return form.fields.map((label, index) => field(label, style, index === 0)).join("") + closing;
}

/** Thân văn bản đầy đủ: khối đầu, tên loại và trích yếu, các mục của mẫu, khối cuối. */
export function assembleSmallFormBody(input: AssembleInput, form: SmallForm, stylePrefix: string): string {
  if (!input.title || input.title.variant !== "named") throw new Error("Giấy tờ này phải dùng khối tên loại và trích yếu.");
  return buildHeaderBlock(input.header) + buildTitleBody(input.title) + buildSmallFormBody(form, stylePrefix) + buildFooterBlock(input.footer);
}
