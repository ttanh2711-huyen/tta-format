import { summarizeIssues } from "./engine";
import type { ReportHeader } from "./exportReport";
import type { Issue } from "./types";
import { zipStore } from "../ooxml/zipStore";

export { crc32, toBase64 } from "../ooxml/zipStore";

/**
 * Báo cáo kiểm tra dạng .docx (T26): tạo gói OOXML tối thiểu rồi đóng zip kiểu "store" (không nén) nên không cần
 * thư viện; Word mở được. Dùng với `Application.createDocument(base64).open()`.
 */

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
  return documentOf(body.join(""), false);
}

/** Thân tài liệu A4 (dọc hoặc ngang). */
function documentOf(bodyXml: string, landscape: boolean): string {
  const size = landscape ? `<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>` : `<w:pgSz w:w="11906" w:h="16838"/>`;
  const sect = `<w:sectPr>${size}<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml}${sect}</w:body></w:document>`;
}

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

function packDocx(documentXmlText: string): Uint8Array {
  const encoder = new TextEncoder();
  return zipStore([
    { name: "[Content_Types].xml", data: encoder.encode(CONTENT_TYPES) },
    { name: "_rels/.rels", data: encoder.encode(ROOT_RELS) },
    { name: "word/document.xml", data: encoder.encode(documentXmlText) },
  ]);
}

export function buildReportDocx(issues: readonly Issue[], header: ReportHeader): Uint8Array {
  return packDocx(documentXml(issues, header));
}

/** Tài liệu Word khổ ngang có tiêu đề và một bảng kẻ ô (dòng đầu đậm), ví dụ sổ đăng ký văn bản đi (T27). */
export function buildTableDocx(title: string, rows: readonly (readonly string[])[]): Uint8Array {
  const border = ["top", "left", "bottom", "right", "insideH", "insideV"].map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join("");
  const cell = (text: string, bold: boolean): string =>
    `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>${bold ? "<w:b/>" : ""}<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:tc>`;
  const table = `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${border}</w:tblBorders><w:tblLayout w:type="autofit"/></w:tblPr>${rows.map((row, r) => `<w:tr>${r === 0 ? "<w:trPr><w:tblHeader/></w:trPr>" : ""}${row.map((text) => cell(text, r === 0)).join("")}</w:tr>`).join("")}</w:tbl>`;
  const heading = `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:b/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">${escapeXml(title)}</w:t></w:r></w:p>`;
  return packDocx(documentOf(`${heading}${table}<w:p/>`, true));
}
