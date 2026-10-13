import { summarizeIssues } from "./engine";
import type { ReportHeader } from "./exportReport";
import type { Issue } from "./types";

/**
 * Báo cáo kiểm tra dạng .docx (T26): tạo gói OOXML tối thiểu rồi đóng zip kiểu "store" (không nén) nên không cần
 * thư viện; Word mở được. Dùng với `Application.createDocument(base64).open()`.
 */

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

export function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff;
  for (const b of bytes) c = CRC_TABLE[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

interface Entry {
  name: string;
  data: Uint8Array;
}

/** Zip không nén: mỗi tệp một local header, rồi central directory và bản ghi kết thúc. */
function zipStore(entries: readonly Entry[]): Uint8Array {
  const encoder = new TextEncoder();
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  const u16 = (view: DataView, at: number, value: number): void => view.setUint16(at, value, true);
  const u32 = (view: DataView, at: number, value: number): void => view.setUint32(at, value, true);
  for (const { name, data } of entries) {
    const nameBytes = encoder.encode(name);
    const crc = crc32(data);
    const local = new Uint8Array(30 + nameBytes.length);
    const lv = new DataView(local.buffer);
    u32(lv, 0, 0x04034b50);
    u16(lv, 4, 20);
    u16(lv, 6, 0x0800); // tên tệp UTF-8
    u32(lv, 14, crc);
    u32(lv, 18, data.length);
    u32(lv, 22, data.length);
    u16(lv, 26, nameBytes.length);
    local.set(nameBytes, 30);
    const entry = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(entry.buffer);
    u32(cv, 0, 0x02014b50);
    u16(cv, 4, 20);
    u16(cv, 6, 20);
    u16(cv, 8, 0x0800);
    u32(cv, 16, crc);
    u32(cv, 20, data.length);
    u32(cv, 24, data.length);
    u16(cv, 28, nameBytes.length);
    u32(cv, 42, offset);
    entry.set(nameBytes, 46);
    central.push(entry);
    chunks.push(local, data);
    offset += local.length + data.length;
  }
  const centralSize = central.reduce((sum, c) => sum + c.length, 0);
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  u32(ev, 0, 0x06054b50);
  u16(ev, 8, entries.length);
  u16(ev, 10, entries.length);
  u32(ev, 12, centralSize);
  u32(ev, 16, offset);
  const all = [...chunks, ...central, end];
  const out = new Uint8Array(all.reduce((sum, c) => sum + c.length, 0));
  let at = 0;
  for (const part of all) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

const escapeXml = (text: string): string => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

function paragraph(text: string, options: { bold?: boolean; sizeHalfPt?: number; indent?: boolean } = {}): string {
  const rPr = `<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>${options.bold ? "<w:b/>" : ""}<w:sz w:val="${options.sizeHalfPt ?? 26}"/></w:rPr>`;
  const pPr = options.indent ? `<w:pPr><w:ind w:left="567"/></w:pPr>` : "";
  return `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function documentXml(issues: readonly Issue[], header: ReportHeader): string {
  const counts = summarizeIssues(issues);
  const fixable = issues.filter((issue) => issue.fix).length;
  const body: string[] = [
    paragraph(`BÁO CÁO KIỂM TRA VĂN BẢN THEO CHUẨN ${header.standardName.toUpperCase()}`, { bold: true, sizeHalfPt: 28 }),
    paragraph(`Kiểm tra lúc: ${header.checkedAt.toLocaleString("vi-VN")}`),
    paragraph(`Tổng: ${counts.errors} lỗi, ${counts.warnings} cảnh báo; ${fixable} mục sửa được tự động.`),
  ];
  if (issues.length === 0) body.push(paragraph("Không thấy lỗi."));
  issues.forEach((issue, i) => {
    const where = issue.location.kind === "page" ? `Trang, phần ${issue.location.section}` : `Đoạn "${issue.location.excerpt}"`;
    body.push(paragraph(`${i + 1}. [${issue.severity === "error" ? "Lỗi" : "Cảnh báo"}] ${issue.message}`, { bold: true }));
    body.push(paragraph(`Vị trí: ${where}`, { indent: true }));
    body.push(paragraph(`Căn cứ: ${issue.reference}`, { indent: true }));
    if (issue.actual !== undefined || issue.expected !== undefined) body.push(paragraph(`Hiện có: ${issue.actual ?? "-"} | Cần: ${issue.expected ?? "-"}`, { indent: true }));
    body.push(paragraph(`Cách sửa: ${issue.fix ? issue.fix.label : "cần sửa thủ công"}`, { indent: true }));
  });
  const sect = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body.join("")}${sect}</w:body></w:document>`;
}

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

export function buildReportDocx(issues: readonly Issue[], header: ReportHeader): Uint8Array {
  const encoder = new TextEncoder();
  return zipStore([
    { name: "[Content_Types].xml", data: encoder.encode(CONTENT_TYPES) },
    { name: "_rels/.rels", data: encoder.encode(ROOT_RELS) },
    { name: "word/document.xml", data: encoder.encode(documentXml(issues, header)) },
  ]);
}

export function toBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}
