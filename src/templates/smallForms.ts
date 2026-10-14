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
  /** Nguồn bố cục của chuẩn Chính quyền, hiện cho người dùng. */
  source: string;
  /** Ghi chú cho chuẩn Đảng (Hướng dẫn 05 không có mẫu riêng cho giấy tờ); bỏ trống thì dùng `source`. */
  sourceParty?: string;
}

const UNVERIFIED = "bố cục theo cách trình bày thông dụng, chưa đối chiếu nguyên văn";
const PARTY_NOTE = "Hướng dẫn 05 không có mẫu riêng cho giấy tờ này: bố cục đề xuất, cần bạn duyệt";

const SAO_NOTE = "Phụ lục III chỉ nêu chữ viết tắt SY, TrS, SL (Mục I); bố cục bản sao là đề xuất, cần bạn duyệt";

const CHI_BO_NOTE = "Bố cục theo cách trình bày thông dụng của văn bản chi bộ, đảng ủy cơ sở (nghị quyết hằng tháng, kiểm điểm cuối năm); là đề xuất, cần đối chiếu hướng dẫn hiện hành của cấp ủy cấp trên";
const PARTY_ONLY = "Chỉ dùng cho chuẩn Đảng";

export const SMALL_FORMS: readonly SmallForm[] = [
  { docTypeId: "giay-moi", fields: ["Kính mời:", "Đến dự:", "Thời gian:", "Địa điểm:", "Nội dung:"], closing: "Rất mong sự có mặt của quý vị.", source: `Mẫu 1.7, ${UNVERIFIED} Phụ lục III` },
  { docTypeId: "giay-gioi-thieu", fields: ["Kính gửi:", "Trân trọng giới thiệu ông (bà):", "Chức vụ:", "Được cử đến:", "Về việc:"], closing: "Đề nghị quý cơ quan tạo điều kiện để ông (bà) có tên trên hoàn thành nhiệm vụ.", source: `Mẫu 1.8, ${UNVERIFIED} Phụ lục III` },
  { docTypeId: "cong-dien", fields: ["Kính gửi:", "Nội dung:"], closing: "", source: `Mẫu 1.6, ${UNVERIFIED} Phụ lục III` },
  { docTypeId: "giay-nghi-phep", fields: ["Cấp cho ông (bà):", "Chức vụ:", "Đơn vị công tác:", "Được nghỉ phép từ ngày … đến ngày …:", "Nơi nghỉ phép:"], closing: "", source: `Mẫu 1.10, ${UNVERIFIED} Phụ lục III` },
  { docTypeId: "thu-cong", fields: ["Kính gửi:", "Nội dung:"], closing: "Trân trọng.", source: "Phụ lục III không có mẫu: bố cục đề xuất, cần bạn duyệt" },
  { docTypeId: "sao-y", fields: ["Sao y bản chính văn bản số … ngày … của …:", "Nội dung:"], closing: "", source: SAO_NOTE },
  { docTypeId: "trich-sao", fields: ["Trích sao từ văn bản số … ngày … của …:", "Nội dung trích:"], closing: "", source: SAO_NOTE },
  { docTypeId: "sao-luc", fields: ["Sao lục từ văn bản số … ngày … của …:", "Nội dung:"], closing: "", source: SAO_NOTE },
  {
    docTypeId: "nghi-quyet-chi-bo",
    fields: [
      "I. Đánh giá kết quả thực hiện nghị quyết tháng trước",
      "1. Kết quả đạt được:",
      "2. Hạn chế, nguyên nhân:",
      "II. Nhiệm vụ, giải pháp tháng tới",
      "1. Về công tác chính trị, tư tưởng:",
      "2. Về thực hiện nhiệm vụ chính trị của cơ quan, đơn vị:",
      "3. Về công tác xây dựng Đảng:",
      "III. Tổ chức thực hiện",
    ],
    closing: "Nghị quyết được chi bộ thông qua tại cuộc họp ngày … tháng … năm …, tán thành …/… đảng viên dự họp.",
    source: PARTY_ONLY,
    sourceParty: CHI_BO_NOTE,
  },
  {
    docTypeId: "bao-cao-kiem-diem-tap-the",
    fields: [
      "I. Ưu điểm, kết quả đạt được",
      "II. Hạn chế, khuyết điểm và nguyên nhân",
      "III. Kết quả khắc phục những hạn chế, khuyết điểm đã được cấp có thẩm quyền kết luận hoặc được chỉ ra ở các kỳ kiểm điểm trước",
      "IV. Giải trình những vấn đề được gợi ý kiểm điểm (nếu có)",
      "V. Làm rõ trách nhiệm của tập thể đối với những hạn chế, khuyết điểm (nếu có)",
      "VI. Phương hướng, biện pháp khắc phục hạn chế, khuyết điểm",
      "VII. Tự nhận mức xếp loại chất lượng",
    ],
    closing: "",
    source: PARTY_ONLY,
    sourceParty: CHI_BO_NOTE,
  },
  {
    docTypeId: "ban-kiem-diem-ca-nhan",
    fields: [
      "Họ và tên:",
      "Ngày sinh:",
      "Chức vụ đảng, chính quyền, đoàn thể:",
      "Đơn vị công tác:",
      "Sinh hoạt tại chi bộ:",
      "I. Ưu điểm, kết quả đạt được",
      "1. Về tư tưởng chính trị:",
      "2. Về phẩm chất đạo đức, lối sống; ý thức tổ chức kỷ luật; tác phong, lề lối làm việc:",
      "3. Về thực hiện chức trách, nhiệm vụ được giao:",
      "4. Việc thực hiện cam kết tu dưỡng, rèn luyện, phấn đấu hằng năm:",
      "II. Hạn chế, khuyết điểm và nguyên nhân",
      "III. Kết quả khắc phục những hạn chế, khuyết điểm đã được cấp có thẩm quyền kết luận hoặc được chỉ ra ở các kỳ kiểm điểm trước",
      "IV. Giải trình những vấn đề được gợi ý kiểm điểm (nếu có)",
      "V. Làm rõ trách nhiệm của cá nhân đối với những hạn chế, khuyết điểm của tập thể (nếu có)",
      "VI. Phương hướng, biện pháp khắc phục hạn chế, khuyết điểm",
      "VII. Tự nhận mức xếp loại chất lượng",
    ],
    closing: "",
    source: PARTY_ONLY,
    sourceParty: CHI_BO_NOTE,
  },
  { docTypeId: "giay-chung-nhan", fields: ["Chứng nhận ông (bà):", "Chức vụ, đơn vị:", "Nội dung chứng nhận:"], closing: "", source: "Chỉ dùng cho chuẩn Đảng", sourceParty: PARTY_NOTE },
  { docTypeId: "giay-di-duong", fields: ["Cấp cho ông (bà):", "Chức vụ, đơn vị:", "Được cử đi công tác tại:", "Thời gian từ ngày … đến ngày …:", "Phương tiện:"], closing: "", source: "Chỉ dùng cho chuẩn Đảng", sourceParty: PARTY_NOTE },
  { docTypeId: "phieu-chuyen", fields: ["Kính gửi:", "Chuyển kèm theo:", "Nội dung đề nghị:"], closing: "", source: "Chỉ dùng cho chuẩn Đảng", sourceParty: PARTY_NOTE },
  { docTypeId: "phieu-gui", fields: ["Kính gửi:", "Gửi kèm theo:", "Nội dung:"], closing: "", source: "Chỉ dùng cho chuẩn Đảng", sourceParty: PARTY_NOTE },
];

/** Nguồn bố cục hiện cho người dùng, theo kiểu khối đầu của chuẩn đang chọn. */
export function formSource(form: SmallForm, headerKind: "quocHieu" | "dang"): string {
  return headerKind === "dang" ? (form.sourceParty ?? PARTY_NOTE) : form.source;
}

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
