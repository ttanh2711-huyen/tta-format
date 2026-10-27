import { STANDARDS } from "../standards";
import { buildHeaderBlock, headerInputFromOptions, QUOC_HIEU, TIEU_NGU } from "../ooxml/headerBlock";
import { buildTitleBody, paragraph, recipientsRowOf, subjectLine, titleInputFromOptions } from "../ooxml/titleBlock";
import type { FormatValues } from "../settings/formatOptions";
import { buildPartyHeaderBlock, partyHeaderFromOptions } from "../standards/hd05/headerBlock";
import { buildPartyTitleBody, dashList, partyRecipientsRow, partySubjectRow, partyTitleFromOptions, type PartyTitleVariant } from "../standards/hd05/titleBlock";
import type { StandardProfile } from "../standards/types";
import { matchTitle, pad2 } from "./common";
import type { CheckDocument, CheckParagraph } from "./types";

/**
 * Sửa khối đầu có sẵn nhưng chưa đúng quy định, phần thuần: đọc nội dung khối đầu cũ (cơ quan, số ký hiệu, địa danh ngày
 * tháng, tên loại, trích yếu, "Kính gửi") ở đầu văn bản, dù đặt trong bảng hay trong đoạn thường, rồi chuẩn hoá theo quy
 * định để dựng lại khối đầu mới. Vùng khối đầu cũ (từ đầu văn bản tới hết trích yếu, "Kính gửi") được xoá thay bằng khối mới.
 */

export interface HeaderFields {
  /** Cơ quan chủ quản (Chính quyền) hoặc cơ quan cấp trên (Đảng). */
  parentLines: string[];
  issuerLines: string[];
  /** Dòng số, ký hiệu đã chuẩn hoá ("Số: 05/BC-UBND" hoặc "Số 05-BC/ĐU"); trống nếu văn bản không có. */
  so: string;
  /** Địa danh, ngày tháng đã chuẩn hoá; trống nếu không có. */
  placeDate: string;
  /** letter: công văn (trích yếu "V/v", "về việc" dưới số); named: văn bản có tên loại. */
  kind: "letter" | "named";
  typeName: string;
  /** Trích yếu, không gồm "V/v", "về việc". */
  subject: string;
  /** Nơi nhận của "Kính gửi" ("Kính trình" với tờ trình Đảng). */
  recipients: string[];
}

export interface HeaderRebuildPlan {
  fields: HeaderFields;
  /** Vùng khối đầu cũ: từ đoạn 0 tới đoạn `end` (chỉ số trong body.paragraphs). */
  end: number;
  /** Các bảng (chỉ số bảng) nằm trong vùng cũ, xoá cả bảng. */
  tables: number[];
  notes: string[];
}

const DATE = /ngày\s*\d*\s*tháng\s*\d*\s*năm/u;
const SO = /^Số(?![\p{L}])/u;
const SUBJECT_LETTER = /^(?:V\/v|v\/v|về việc|Về việc)(?![\p{L}])\s*:?\s*/u;
const KINH_GUI = /^Kính\s+(gửi|trình)\s*:?\s*/u;
const RULE_LINE = /^[\s\-–—_*.=]+$/u;
const ADMIN_TAIL = /^(XÃ|PHƯỜNG|THỊ TRẤN|THỊ XÃ|HUYỆN|QUẬN|THÀNH PHỐ|TỈNH|ĐẶC KHU)(?![\p{L}])/u;
const ADMIN_HEAD = /^(ỦY BAN NHÂN DÂN|UỶ BAN NHÂN DÂN|UBND|HỘI ĐỒNG NHÂN DÂN|HĐND)$/u;
const BODY_START = /^(Căn cứ|Xét|Theo đề nghị|Thực hiện|Kính trình)(?![\p{L}])/u;

const clean = (s: string): string => s.replace(/\s+/g, " ").trim();
const text = (p: CheckParagraph): string => clean(p.resolved.text);
const isUpper = (s: string): boolean => {
  const letters = s.replace(/[^\p{L}]/gu, "");
  return letters.length > 0 && letters === letters.toUpperCase();
};

/** Tiêu đề, tiêu ngữ của bất kỳ chuẩn nào (khối mới sẽ dựng lại đúng chuẩn đang chọn). */
function isTitleLine(s: string): boolean {
  const titles = [QUOC_HIEU, TIEU_NGU, ...STANDARDS.map((standard) => standard.identifyingTitle)];
  return titles.some((title) => matchTitle(s, title) !== "no") || /^Độc lập\s*[-–—]?\s*Tự do/iu.test(s);
}

/** Tên loại văn bản chuẩn ("Báo cáo") nếu dòng in hoa là một tên loại đã biết, không thì chính dòng đó. */
function canonicalType(s: string): string {
  const upper = clean(s).toUpperCase();
  for (const standard of STANDARDS) {
    const found = standard.docTypes.find((type) => type.name.toUpperCase() === upper);
    if (found) return found.name;
  }
  return clean(s);
}

const knownType = (s: string): boolean => STANDARDS.some((standard) => standard.docTypes.some((type) => type.name.toUpperCase() === clean(s).toUpperCase()));

/** "Số: 5 / BC - UBND" thành "Số: 05/BC-UBND" (Chính quyền) hoặc "Số 05-BC/ĐU" (Đảng); để trống số thì giữ chỗ. */
export function normalizeSo(raw: string, profile: StandardProfile): string {
  const rest = clean(raw).replace(/^Số\s*[:.]?\s*/u, "");
  const squeeze = (s: string): string => s.replace(/\s*-\s*/g, "-").replace(/\s*\/\s*/g, "/").replace(/\s+/g, "");
  if (profile.headerKind === "dang") {
    const m = /^(\d*)\s*-\s*(.*)$/u.exec(rest);
    if (m) return m[1] === "" ? `Số    -${squeeze(m[2])}` : `Số ${pad2(m[1])}-${squeeze(m[2])}`;
    return `Số ${squeeze(rest)}`;
  }
  const m = /^(\d*)\s*\/\s*(.*)$/u.exec(rest);
  if (m) return m[1] === "" ? `Số:      /${squeeze(m[2])}` : `Số: ${pad2(m[1])}/${squeeze(m[2])}`;
  return `Số: ${squeeze(rest)}`;
}

/** "Hà Nội ngày 5 tháng 3 năm 2026" thành "Hà Nội, ngày 05 tháng 3 năm 2026": ngày dưới 10 và tháng 1, 2 thêm số 0. */
export function normalizeDate(raw: string): string {
  const m = /^(.*?)[,\s]*ngày\s*(\d*)\s*tháng\s*(\d*)\s*năm\s*(\d*)\s*\.?$/u.exec(clean(raw));
  if (!m) return clean(raw);
  const [, place, day, month, year] = m;
  const blank = "    ";
  const d = day === "" ? blank : pad2(day);
  const mo = month === "" ? blank : Number(month) < 3 ? pad2(String(Number(month))) : String(Number(month));
  return `${clean(place)}, ngày ${d} tháng ${mo} năm ${year === "" ? blank : year}`;
}

/** Nơi nhận: tách theo dòng hoặc dấu chấm phẩy, bỏ gạch đầu dòng và dấu câu cuối. */
function recipientItems(lines: readonly string[]): string[] {
  return lines
    .flatMap((line) => line.split(/;\s*(?=\S)/))
    .map((item) => clean(item).replace(/^[-–—+•]\s*/u, "").replace(/[;.,]+$/u, "").trim())
    .filter((item) => item !== "");
}

/** Viết hoa chữ đầu trích yếu, bỏ dấu chấm cuối. */
function tidySubject(s: string): string {
  const t = clean(s).replace(/[.;]+$/u, "");
  return t.charAt(0).toUpperCase() + t.slice(1);
}

/** Tách cơ quan chủ quản và cơ quan ban hành: theo in đậm nếu có; không thì dòng cuối là cơ quan ban hành. */
function splitAgency(lines: readonly CheckParagraph[], notes: string[]): { parentLines: string[]; issuerLines: string[] } {
  const texts = lines.map((p) => text(p).toUpperCase());
  if (texts.length === 0) return { parentLines: [], issuerLines: [] };
  const firstBold = lines.findIndex((p) => p.resolved.bold === true);
  if (firstBold > 0 && lines.slice(firstBold).every((p) => p.resolved.bold === true)) {
    return { parentLines: texts.slice(0, firstBold), issuerLines: texts.slice(firstBold) };
  }
  if (texts.length === 1) return { parentLines: [], issuerLines: texts };
  // "ỦY BAN NHÂN DÂN" / "XÃ A" là một tên cơ quan xuống dòng.
  if (texts.length === 2 && ADMIN_HEAD.test(texts[0]) && ADMIN_TAIL.test(texts[1])) return { parentLines: [], issuerLines: [texts.join(" ")] };
  notes.push("Không phân biệt chắc cơ quan chủ quản và cơ quan ban hành (cùng kiểu chữ); đã lấy dòng cuối làm cơ quan ban hành, hãy xem lại.");
  return { parentLines: texts.slice(0, -1), issuerLines: texts.slice(-1) };
}

/** Thiếu gì để dựng được khối đầu mới. */
export function validateHeaderFields(fields: HeaderFields): string[] {
  const errors: string[] = [];
  if (fields.issuerLines.every((line) => clean(line) === "")) errors.push("Hãy nhập tên cơ quan ban hành.");
  if (fields.kind === "named" && clean(fields.typeName) === "") errors.push("Hãy nhập tên loại văn bản.");
  if (clean(fields.subject) === "") errors.push("Hãy nhập trích yếu nội dung.");
  return errors;
}

/**
 * XML khối đầu mới đúng quy định (bảng khối đầu, tên loại, trích yếu, "Kính gửi") dựng bằng chính các bộ dựng của trình
 * tạo văn bản, theo chuẩn đang chọn và Tuỳ chọn nhanh (cỡ chữ, lề).
 */
export function buildRebuiltHeader(fields: HeaderFields, profile: StandardProfile, values: FormatValues): string {
  const problems = validateHeaderFields(fields);
  if (problems.length > 0) throw new Error(problems.join(" "));
  const parentLines = fields.parentLines.map(clean).filter((line) => line !== "");
  const issuerLines = fields.issuerLines.map(clean).filter((line) => line !== "");
  const recipients = fields.recipients.map(clean).filter((line) => line !== "");
  const numberRow = fields.so !== "" || fields.placeDate !== "" ? { so: fields.so, placeDate: fields.placeDate } : undefined;
  const letter = fields.kind === "letter";

  if (profile.headerKind === "dang") {
    const variant: PartyTitleVariant = letter ? "congVan" : /^tờ trình$/iu.test(clean(fields.typeName)) && recipients.length > 0 ? "toTrinh" : "coTenLoai";
    const title = partyTitleFromOptions(profile, values, { variant, typeName: fields.typeName, subject: fields.subject, recipients, basis: [] });
    const header = partyHeaderFromOptions(profile, values, {
      parentLines,
      issuerLines,
      numberRow,
      subjectRow: partySubjectRow(title),
      recipientsRow: letter && recipients.length > 0 ? partyRecipientsRow(title) : undefined,
    });
    return buildPartyHeaderBlock(header) + (letter ? "" : buildPartyTitleBody(title));
  }

  const title = titleInputFromOptions(profile, values, { variant: letter ? "letter" : "named", typeName: fields.typeName, subject: fields.subject, recipients });
  const header = headerInputFromOptions(profile, values, {
    parentLines,
    issuer: issuerLines.join(" "),
    numberRow,
    subjectRow: letter ? { text: subjectLine(fields.subject) } : undefined,
    recipientsRow: letter && recipients.length > 0 ? recipientsRowOf(title) : undefined,
  });
  let xml = buildHeaderBlock(header) + (letter ? "" : buildTitleBody(title));
  // Văn bản có tên loại mà có "Kính gửi" (tờ trình): dòng "Kính gửi" căn giữa dưới trích yếu.
  if (!letter && recipients.length > 0) {
    const size = title.sizes.kinhGui;
    xml +=
      recipients.length === 1
        ? paragraph(`Kính gửi: ${recipients[0]}`, { sizePt: size, beforePt: 6, afterPt: 6 })
        : paragraph("Kính gửi:", { sizePt: size, beforePt: 6 }) + dashList(recipients).map((line) => paragraph(line, { sizePt: size, align: "left", leftMm: 0, hangingMm: 4 })).join("");
  }
  return xml;
}

export function planHeaderRebuild(doc: CheckDocument, profile: StandardProfile): HeaderRebuildPlan | { error: string } {
  const paras = doc.paragraphs;
  const notes: string[] = [];
  // Vùng xét: các đoạn ở đầu văn bản trước thân; bảng khối đầu là bảng đầu tiên nếu có tiêu đề hoặc dòng số trong đó.
  const firstText = paras.find((p) => text(p) !== "");
  if (!firstText) return { error: "Văn bản trống." };
  const headTable = firstText.table && paras.some((p) => p.table?.index === firstText.table?.index && (isTitleLine(text(p)) || SO.test(text(p)) || DATE.test(text(p)))) ? firstText.table.index : undefined;

  const agency: CheckParagraph[] = [];
  let so = "";
  let placeDate = "";
  let subjectLines: string[] = [];
  let kind: "letter" | "named" | undefined;
  let typeName = "";
  const recipientLines: string[] = [];
  let sawTitle = false;
  let end = -1;
  const tables = new Set<number>();
  /** Đang gom tiếp dòng của mục nào (trích yếu công văn, Kính gửi). */
  let collecting: "subject" | "recipients" | null = null;

  for (const p of paras) {
    const t = text(p);
    const inHeadTable = headTable !== undefined && p.table?.index === headTable;
    if (p.table && !inHeadTable) break; // bảng khác: hết vùng khối đầu
    if (inHeadTable) tables.add(headTable);
    if (t === "") {
      if (inHeadTable) end = p.index;
      if (!inHeadTable) collecting = collecting === "recipients" ? collecting : null;
      continue;
    }
    if (isTitleLine(t)) {
      sawTitle = true;
      end = p.index;
      collecting = null;
      continue;
    }
    if (RULE_LINE.test(t)) {
      end = p.index;
      collecting = null;
      continue;
    }
    if (SO.test(t) && so === "") {
      so = normalizeSo(t, profile);
      end = p.index;
      collecting = null;
      continue;
    }
    if (DATE.test(t) && placeDate === "") {
      placeDate = normalizeDate(t);
      end = p.index;
      collecting = null;
      continue;
    }
    if (SUBJECT_LETTER.test(t) && kind === undefined) {
      kind = "letter";
      subjectLines = [t.replace(SUBJECT_LETTER, "")];
      collecting = "subject";
      end = p.index;
      continue;
    }
    const kg = KINH_GUI.exec(t);
    if (kg) {
      const rest = t.slice(kg[0].length);
      if (rest !== "") recipientLines.push(rest);
      collecting = "recipients";
      end = p.index;
      continue;
    }
    if (collecting === "recipients" && (/^[-–—+•]/u.test(t) || inHeadTable)) {
      recipientLines.push(t);
      end = p.index;
      continue;
    }
    if (inHeadTable) {
      // Trong bảng khối đầu: dòng nối tiếp của trích yếu công văn, hoặc tên cơ quan (cột trái, hàng đầu).
      if (collecting === "subject") subjectLines.push(t);
      else if (p.table?.col === 0 && p.table.row === 0) agency.push(p);
      end = p.index;
      continue;
    }
    if (headTable === undefined && !sawTitle && so === "" && kind === undefined && isUpper(t) && agency.length < 5) {
      // Khối đầu không dựng bằng bảng: tên cơ quan là các dòng in hoa trên cùng, trước quốc hiệu và dòng số.
      agency.push(p);
      end = p.index;
      continue;
    }
    // Sau bảng khối đầu: tên loại in hoa, rồi trích yếu (căn giữa hoặc đậm, tối đa 3 dòng).
    if (kind === undefined && isUpper(t) && t.length <= 80 && !BODY_START.test(t) && (knownType(t) || p.resolved.align === "center" || p.resolved.bold === true)) {
      kind = "named";
      typeName = canonicalType(t);
      collecting = "subject";
      end = p.index;
      continue;
    }
    if (collecting === "subject" && kind === "named" && subjectLines.length < 3 && !BODY_START.test(t) && (p.resolved.align === "center" || p.resolved.bold === true)) {
      subjectLines.push(t);
      end = p.index;
      continue;
    }
    break;
  }

  const { parentLines, issuerLines } = splitAgency(agency, notes);
  if (end < 0 || (issuerLines.length === 0 && so === "" && !sawTitle)) return { error: "Không nhận ra khối đầu ở đầu văn bản (không thấy tên cơ quan, quốc hiệu hay dòng số)." };
  if (issuerLines.length === 0) notes.push("Không thấy tên cơ quan ban hành; hãy nhập trước khi dựng lại.");
  if (kind === undefined) notes.push("Không thấy tên loại hay trích yếu; mặc định là công văn, hãy kiểm tra.");
  const subject = subjectLines.length > 0 ? tidySubject(subjectLines.join(" ")) : "";
  if (subject === "") notes.push("Không thấy trích yếu nội dung; hãy nhập trước khi dựng lại.");
  return {
    fields: { parentLines, issuerLines, so, placeDate, kind: kind ?? "letter", typeName, subject, recipients: recipientItems(recipientLines) },
    end,
    tables: [...tables],
    notes,
  };
}
