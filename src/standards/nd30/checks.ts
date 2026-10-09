import { checkDate, checkTitleParagraph, pad2, paragraphIssue, replaceFix } from "../../checker/common";
import type { Fix, Issue, CheckDocument, CheckParagraph } from "../../checker/types";
import { findAcademicTitle } from "../../ooxml/footerBlock";
import { QUOC_HIEU, TIEU_NGU } from "../../ooxml/headerBlock";
import type { StandardProfile } from "../types";

/**
 * Quy tắc kiểm tra khối đầu và khối cuối của chuẩn Chính quyền (T14): quốc hiệu, tiêu ngữ, số và ký hiệu, ngày
 * tháng, "Nơi nhận", chức vụ, quyền hạn, họ tên người ký. Chỉ xét các đoạn nằm trong bảng (khối đầu và khối cuối
 * dựng bằng bảng) nên không đụng tới phần thân.
 */

const AUTHORITY_WITHOUT_DOT = /^(TM|KT|TL|TUQ|Q)(?=\s)/;

/** "Số: 05/QĐ-UBND": có dấu hai chấm, số < 10 thêm số 0, không khoảng trắng quanh "/" và "-". */
function checkSo(p: CheckParagraph, source: string): Issue[] {
  const raw = p.resolved.text.trim();
  const head = /^Số\s*(:)?/.exec(raw);
  if (!head) return [];
  const m = /^(\s*)(\d*)(\s*)\/(\s*)(.*)$/.exec(raw.slice(head[0].length));
  if (!m) return [];
  const [, pre, number, post, afterSlash, rest] = m;
  const problems: string[] = [];
  if (!head[1]) problems.push("thiếu dấu hai chấm sau \"Số\"");
  const blank = number === "";
  if (!blank) {
    if (number.length === 1) problems.push("số nhỏ hơn 10 phải thêm số 0 phía trước");
    if (pre === "") problems.push("thiếu khoảng trắng sau \"Số:\"");
    else if (pre.length > 1) problems.push("thừa khoảng trắng sau \"Số:\"");
    if (post !== "" || afterSlash !== "") problems.push("thừa khoảng trắng quanh dấu \"/\"");
  }
  const cleanRest = rest.replace(/\s*-\s*/g, "-").replace(/\s+/g, "");
  if (cleanRest !== rest) problems.push("ký hiệu không được có khoảng trắng, các phần nối bằng \"-\" không cách chữ");
  if (problems.length === 0) return [];
  const expected = blank ? `Số:${pre}${post}/${cleanRest}` : `Số: ${pad2(number)}/${cleanRest}`;
  return [
    paragraphIssue(p, {
      ruleId: "header.so.format",
      severity: "error",
      message: `Số và ký hiệu "${raw}" chưa đúng: ${problems.join("; ")}.`,
      reference: source,
      actual: raw,
      expected,
      fix: replaceFix(p, expected, "Sửa số và ký hiệu"),
    }),
  ];
}


function cellKey(p: CheckParagraph): string {
  return p.table ? `${p.table.index}:${p.table.row}:${p.table.col}` : "";
}

/** Khối cuối: "Nơi nhận" và danh sách, chức vụ, quyền hạn, họ tên người ký. */
function checkFooter(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const issues: Issue[] = [];
  const cells = new Map<string, CheckParagraph[]>();
  for (const p of doc.paragraphs) {
    if (!p.table || p.resolved.text.trim() === "") continue;
    const key = cellKey(p);
    cells.set(key, [...(cells.get(key) ?? []), p]);
  }
  const noiNhanTitle = profile.elements.noiNhanTieuDe;
  const noiNhanList = profile.elements.noiNhanDanhSach;
  const chucVu = profile.elements.quyenHanChucVu;
  const hoTen = profile.elements.hoTenNguoiKy;

  for (const [key, cell] of cells) {
    const headingAt = cell.findIndex((p) => /^nơi nhận\s*:$/i.test(p.resolved.text.trim()));
    if (headingAt < 0) continue;
    const heading = cell[headingAt];
    const fontFix = (p: CheckParagraph, patch: { bold?: boolean; italic?: boolean; sizePt?: number }, label: string): Fix => ({ kind: "font", index: p.index, ...patch, label });

    const wrongTitle: string[] = [];
    if (heading.resolved.bold !== true) wrongTitle.push("đậm");
    if (heading.resolved.italic !== true) wrongTitle.push("nghiêng");
    const titleSize = heading.resolved.sizePt;
    const sizeBad = titleSize !== undefined && (titleSize < noiNhanTitle.sizePt.min || titleSize > noiNhanTitle.sizePt.max);
    if (wrongTitle.length > 0 || sizeBad) {
      issues.push(
        paragraphIssue(heading, {
          ruleId: "footer.noiNhan.format",
          severity: "error",
          message: `Từ "Nơi nhận" phải in ${noiNhanTitle.bold ? "đậm, " : ""}nghiêng, cỡ ${noiNhanTitle.sizePt.min} pt.`,
          reference: noiNhanTitle.source,
          fix: fontFix(heading, { bold: true, italic: true, sizePt: noiNhanTitle.sizePt.min }, "Đặt định dạng cho \"Nơi nhận\""),
        }),
      );
    }

    const items = cell.slice(headingAt + 1);
    items.forEach((item, i) => {
      const last = i === items.length - 1;
      const text = item.resolved.text.trim();
      const wanted = last ? "." : ";";
      if (!text.endsWith(wanted)) {
        const expected = `${text.replace(/[;.,]+$/, "")}${wanted}`;
        issues.push(
          paragraphIssue(item, {
            ruleId: "footer.noiNhan.punctuation",
            severity: "error",
            message: last ? "Mục cuối của nơi nhận phải kết thúc bằng dấu chấm." : "Mỗi mục của nơi nhận phải kết thúc bằng dấu chấm phẩy.",
            reference: noiNhanList.source,
            actual: text.slice(-1),
            expected: wanted,
            fix: replaceFix(item, expected, last ? "Thêm dấu chấm cuối mục" : "Thêm dấu chấm phẩy cuối mục"),
          }),
        );
      }
      if (!text.startsWith("-")) {
        issues.push(paragraphIssue(item, { ruleId: "footer.noiNhan.dash", severity: "warning", message: "Mỗi mục của nơi nhận bắt đầu bằng dấu gạch ngang.", reference: noiNhanList.source }));
      }
      const size = item.resolved.sizePt;
      if (size !== undefined && (size < noiNhanList.sizePt.min || size > noiNhanList.sizePt.max)) {
        issues.push(
          paragraphIssue(item, {
            ruleId: "footer.noiNhan.size",
            severity: "error",
            message: `Danh sách nơi nhận cỡ ${size} pt; phải ${noiNhanList.sizePt.min} pt.`,
            reference: noiNhanList.source,
            actual: `${size} pt`,
            expected: `${noiNhanList.sizePt.min} pt`,
            fix: fontFix(item, { sizePt: noiNhanList.sizePt.min }, `Đặt cỡ chữ ${noiNhanList.sizePt.min} pt`),
          }),
        );
      }
    });

    // Ô chữ ký: ô khác cùng hàng của bảng chứa "Nơi nhận".
    const [tableIndex, row, col] = key.split(":").map(Number);
    const signCell = [...cells.entries()].find(([k]) => {
      const [t, r, c] = k.split(":").map(Number);
      return t === tableIndex && r === row && c !== col;
    })?.[1];
    if (!signCell || signCell.length < 2) continue;
    const name = signCell[signCell.length - 1];
    const positions = signCell.slice(0, -1);

    positions.forEach((line, i) => {
      const text = line.resolved.text.trim();
      if (text !== text.toUpperCase()) {
        issues.push(
          paragraphIssue(line, {
            ruleId: "footer.chucVu.upper",
            severity: "error",
            message: "Quyền hạn và chức vụ của người ký phải in hoa.",
            reference: chucVu.source,
            actual: text,
            expected: text.toUpperCase(),
            fix: replaceFix(line, text.toUpperCase(), "In hoa chức vụ"),
          }),
        );
      }
      if (line.resolved.bold !== true) {
        issues.push(
          paragraphIssue(line, {
            ruleId: "footer.chucVu.bold",
            severity: "error",
            message: "Quyền hạn và chức vụ của người ký phải in đậm.",
            reference: chucVu.source,
            fix: fontFix(line, { bold: true }, "In đậm chức vụ"),
          }),
        );
      }
      const size = line.resolved.sizePt;
      if (size !== undefined && (size < chucVu.sizePt.min || size > chucVu.sizePt.max)) {
        issues.push(
          paragraphIssue(line, {
            ruleId: "footer.chucVu.size",
            severity: "error",
            message: `Chức vụ cỡ ${size} pt; phải từ ${chucVu.sizePt.min} đến ${chucVu.sizePt.max} pt.`,
            reference: chucVu.source,
            fix: fontFix(line, { sizePt: chucVu.sizePt.max }, `Đặt cỡ chữ ${chucVu.sizePt.max} pt`),
          }),
        );
      }
      if (i === 0) {
        const match = AUTHORITY_WITHOUT_DOT.exec(text);
        if (match) {
          const expected = `${match[1]}.${text.slice(match[1].length)}`;
          issues.push(
            paragraphIssue(line, {
              ruleId: "footer.quyenHan.format",
              severity: "error",
              message: `Quyền hạn "${match[1]}" phải viết "${match[1]}." (có dấu chấm).`,
              reference: chucVu.source,
              actual: match[1],
              expected: `${match[1]}.`,
              fix: replaceFix(line, expected, "Thêm dấu chấm sau quyền hạn"),
            }),
          );
        }
      }
    });

    const nameText = name.resolved.text.trim();
    const academic = findAcademicTitle(nameText);
    if (academic) {
      issues.push(paragraphIssue(name, { ruleId: "footer.hoTen.hocHam", severity: "error", message: `Họ tên người ký có "${academic}"; chỉ ghi họ và tên, không ghi học hàm, học vị.`, reference: hoTen.source, actual: academic }));
    }
    if (nameText.length >= 3 && nameText === nameText.toUpperCase() && nameText !== nameText.toLowerCase()) {
      issues.push(paragraphIssue(name, { ruleId: "footer.hoTen.case", severity: "error", message: "Họ và tên người ký phải in thường (viết hoa chữ cái đầu).", reference: hoTen.source, actual: nameText }));
    }
    if (name.resolved.bold !== true) {
      issues.push(paragraphIssue(name, { ruleId: "footer.hoTen.bold", severity: "error", message: "Họ và tên người ký phải in đậm.", reference: hoTen.source, fix: fontFix(name, { bold: true }, "In đậm họ tên") }));
    }
  }
  return issues;
}

/** Quy tắc khối đầu và khối cuối của chuẩn Chính quyền, dùng làm `extraRules` của engine. */
export function checkNd30HeaderFooter(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const issues: Issue[] = [];
  const quocHieu = profile.elements.quocHieu;
  const tieuNgu = profile.elements.tieuNgu;
  for (const p of doc.paragraphs) {
    if (!p.inTable || p.resolved.text.trim() === "") continue;
    issues.push(...checkTitleParagraph(p, QUOC_HIEU, "header.quocHieu", "Quốc hiệu", quocHieu.sizePt, quocHieu.source));
    issues.push(...checkTitleParagraph(p, TIEU_NGU, "header.tieuNgu", "Tiêu ngữ", tieuNgu.sizePt, tieuNgu.source));
    issues.push(...checkSo(p, profile.elements.so.source));
    issues.push(...checkDate(p, profile.elements.diaDanhNgayThang.source));
  }
  issues.push(...checkFooter(doc, profile));
  return issues;
}
