import type { StandardProfile } from "../standards/types";
import { matchTitle } from "./common";
import type { CheckDocument, CheckParagraph } from "./types";

/**
 * "Biến văn bản này thành mẫu" (issue #2), phần thuần: tìm các vùng thay đổi theo từng văn bản (số, ngày tháng, tên loại,
 * trích yếu, căn cứ, Kính gửi, thân, nơi nhận, quyền hạn chức vụ, họ tên) để thay bằng Content Control có tiêu đề.
 * Nhận diện dựa trên bố cục thể thức mà bộ kiểm tra đã dùng: bảng khối đầu chứa quốc hiệu hoặc tiêu đề Đảng, bảng khối
 * cuối chứa "Nơi nhận". Vùng nào không nhận ra chắc chắn thì bỏ qua, không đoán.
 */

export type TemplateTag = "so" | "diaDanhNgay" | "trichYeu" | "kinhGui" | "tenLoai" | "canCu" | "noiDung" | "noiNhan" | "chucVu" | "hoTen";

export interface TemplateField {
  tag: TemplateTag;
  /** Tiêu đề hiện trên Content Control. */
  title: string;
  /** Chỉ số đoạn đầu và cuối của vùng (trong `body.paragraphs`). */
  start: number;
  end: number;
  /** Chữ thay vào vùng; null thì giữ nguyên chữ (ví dụ tên loại của mẫu cùng loại). */
  placeholder: string | null;
}

export interface TemplatePlan {
  fields: TemplateField[];
  notes: string[];
}

const TITLES: Record<TemplateTag, string> = {
  so: "Số, ký hiệu văn bản",
  diaDanhNgay: "Địa danh, ngày tháng năm",
  tenLoai: "Tên loại văn bản",
  trichYeu: "Trích yếu nội dung",
  kinhGui: "Kính gửi",
  canCu: "Căn cứ",
  noiDung: "Nội dung văn bản",
  noiNhan: "Nơi nhận",
  chucVu: "Quyền hạn, chức vụ người ký",
  hoTen: "Họ tên người ký",
};

const DATE = /ngày\s.*tháng\s.*năm/u;
const text = (p: CheckParagraph): string => p.resolved.text.trim();
const isUpper = (s: string): boolean => {
  const letters = s.replace(/[^\p{L}]/gu, "");
  return letters.length > 0 && letters === letters.toUpperCase();
};

function field(tag: TemplateTag, start: number, end: number, placeholder: string | null): TemplateField {
  return { tag, title: TITLES[tag], start, end, placeholder };
}

export function planTemplate(doc: CheckDocument, profile: StandardProfile): TemplatePlan {
  const fields: TemplateField[] = [];
  const notes: string[] = [];
  const paras = doc.paragraphs;
  const headerTable = paras.find((p) => p.table && matchTitle(text(p), profile.identifyingTitle) !== "no")?.table?.index;
  if (headerTable === undefined) return { fields, notes: ["Không thấy khối đầu (bảng có quốc hiệu hoặc tiêu đề Đảng) nên không tạo mẫu."] };
  const inHeader = paras.filter((p) => p.table?.index === headerTable && text(p) !== "");
  const footerTable = [...paras].reverse().find((p) => p.table && p.table.index !== headerTable && /^Nơi nhận/u.test(text(p)))?.table?.index;

  // Khối đầu: số, ngày tháng, trích yếu công văn, Kính gửi.
  const so = inHeader.find((p) => /^Số(?!\p{L})/u.test(text(p)));
  if (so) fields.push(field("so", so.index, so.index, text(so).startsWith("Số:") ? "Số: …/…" : "Số …-…/…"));
  const date = inHeader.find((p) => DATE.test(text(p)));
  if (date) fields.push(field("diaDanhNgay", date.index, date.index, "…, ngày … tháng … năm …"));
  const vv = inHeader.find((p) => /^(V\/v|về việc)\s/iu.test(text(p)));
  if (vv) fields.push(field("trichYeu", vv.index, vv.index, text(vv).startsWith("V/v") ? "V/v …" : "về việc …"));
  const kinhGuiLabel = inHeader.find((p) => /^Kính (gửi|trình)\s*:/u.test(text(p)));
  if (kinhGuiLabel?.table) {
    const { row, col } = kinhGuiLabel.table;
    const list = inHeader.filter((p) => p.table?.row === row && p.table.col !== col);
    if (list.length > 0) fields.push(field("kinhGui", list[0].index, list[list.length - 1].index, "…"));
  }

  // Sau bảng khối đầu: tên loại, trích yếu, dòng gạch nối, thẩm quyền, căn cứ, dòng chốt; rồi thân tới bảng khối cuối.
  const headerEnd = Math.max(...paras.filter((p) => p.table?.index === headerTable).map((p) => p.index));
  const footerStart = footerTable === undefined ? paras.length : Math.min(...paras.filter((p) => p.table?.index === footerTable).map((p) => p.index));
  const between = paras.filter((p) => p.index > headerEnd && p.index < footerStart);
  const flow = between.filter((p) => text(p) !== "");
  let cursor = 0;
  const centered = (p: CheckParagraph | undefined): boolean => p !== undefined && !p.inTable && p.resolved.align === "center";
  if (centered(flow[cursor]) && isUpper(text(flow[cursor]))) {
    fields.push(field("tenLoai", flow[cursor].index, flow[cursor].index, null));
    cursor += 1;
    if (centered(flow[cursor]) && !/^-{3,}$/.test(text(flow[cursor])) && !isUpper(text(flow[cursor]))) {
      fields.push(field("trichYeu", flow[cursor].index, flow[cursor].index, profile.headerKind === "dang" ? "…" : "Về việc …"));
      cursor += 1;
    }
  }
  // Dòng gạch nối, thẩm quyền ban hành (in hoa, căn giữa) giữ nguyên.
  while (centered(flow[cursor]) && (/^-{3,}$/.test(text(flow[cursor])) || isUpper(text(flow[cursor])))) cursor += 1;
  const basisStart = cursor;
  while (flow[cursor] && !flow[cursor].inTable && /^(Căn cứ|Xét|Theo đề nghị)(?!\p{L})/u.test(text(flow[cursor]))) cursor += 1;
  if (cursor > basisStart) fields.push(field("canCu", flow[basisStart].index, flow[cursor - 1].index, "Căn cứ …;"));
  // Dòng chốt "QUYẾT ĐỊNH:", "NGHỊ QUYẾT:" giữ nguyên.
  while (centered(flow[cursor]) && isUpper(text(flow[cursor]))) cursor += 1;

  const body = flow.slice(cursor);
  if (body.length > 0) {
    if (body.some((p) => p.inTable)) {
      notes.push("Thân văn bản có bảng nên giữ nguyên thân (không gộp thành một vùng) để không làm hỏng bảng.");
    } else {
      fields.push(field("noiDung", body[0].index, body[body.length - 1].index, "Nội dung văn bản"));
    }
  }

  // Khối cuối: danh sách nơi nhận (cột trái, sau "Nơi nhận:"), quyền hạn chức vụ và họ tên (cột phải).
  if (footerTable !== undefined) {
    const cells = paras.filter((p) => p.table?.index === footerTable);
    const label = cells.find((p) => /^Nơi nhận/u.test(text(p)));
    const left = cells.filter((p) => label?.table && p.table?.col === label.table.col && p.index > label.index && text(p) !== "");
    if (left.length > 0) fields.push(field("noiNhan", left[0].index, left[left.length - 1].index, "- …;"));
    const right = cells.filter((p) => label?.table && p.table?.col !== label.table.col && text(p) !== "");
    if (right.length >= 2) {
      const name = right[right.length - 1];
      const positions = right.slice(0, -1).filter((p) => isUpper(text(p)));
      if (positions.length > 0) fields.push(field("chucVu", positions[0].index, positions[positions.length - 1].index, "QUYỀN HẠN, CHỨC VỤ NGƯỜI KÝ"));
      fields.push(field("hoTen", name.index, name.index, "Họ và tên"));
    }
  } else {
    notes.push("Không thấy khối cuối (bảng có \"Nơi nhận\") nên không tạo vùng nơi nhận, người ký.");
  }
  return { fields: fields.sort((a, b) => a.start - b.start), notes };
}
