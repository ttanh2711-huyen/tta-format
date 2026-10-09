import { checkDate, checkTitleParagraph, DATE_PATTERN, matchTitle, pad2, paragraphIssue, replaceFix } from "../../checker/common";
import type { CheckDocument, CheckParagraph, Fix, Issue } from "../../checker/types";
import { findAcademicTitle } from "../../ooxml/footerBlock";
import { TIEU_NGU } from "../../ooxml/headerBlock";
import type { StandardProfile } from "../types";
import { HEADER_MARKS } from "./rules";

/**
 * Quy tắc kiểm tra khối đầu, tiêu đề và khối cuối của chuẩn Đảng (T14B): tiêu đề "ĐẢNG CỘNG SẢN VIỆT NAM", cơ
 * quan ban hành và dấu *, "Số 05-HD/VPTW", ngày tháng, thể loại, trích yếu và 5 dấu gạch nối, "Kính gửi",
 * "Nơi nhận" gạch chân, quyền hạn "T/M", "K/T", chức vụ không đậm, họ tên người ký.
 */

const AUTHORITY_ABBR = /^(T\/M|K\/T|T\/L|Q\.)(?=\s|$)/;
/** Quyền hạn theo cách viết của Chính quyền và cách viết tương ứng của Đảng. */
const AUTHORITY_GOV: Record<string, string> = { TM: "T/M", KT: "K/T", TL: "T/L" };
const FIRST_TITLE_PARAGRAPHS = 6;
const CONGRESS_NAME = /^ĐẠI HỘI(?=\s|$)/;

type Range = { min: number; max: number };
type FontPatch = Omit<Extract<Fix, { kind: "font" }>, "kind" | "index" | "label">;

const fontFix = (p: CheckParagraph, patch: FontPatch, label: string): Fix => ({ kind: "font", index: p.index, ...patch, label });
const clamp = (n: number, range: Range) => Math.min(range.max, Math.max(range.min, n));
const inRange = (n: number | undefined, range: Range) => n === undefined || (n >= range.min && n <= range.max);
const sizeFix = (p: CheckParagraph, range: Range): Fix => fontFix(p, { sizePt: clamp(p.resolved.sizePt ?? range.min, range) }, `Đặt cỡ chữ ${clamp(p.resolved.sizePt ?? range.min, range)} pt`);

/** "Số 05-HD/VPTW": không dấu hai chấm, số nhỏ hơn 10 thêm số 0, "-" và "/" không có khoảng trắng. */
function checkSo(p: CheckParagraph, source: string): Issue[] {
  const raw = p.resolved.text.trim();
  const head = /^Số(\s*:)?\s*/.exec(raw);
  if (!head) return [];
  const rest = raw.slice(head[0].length);
  const parts = /^(\d*)\s*-\s*([^/]*?)\s*\/\s*(.*)$/.exec(rest);
  if (!parts) {
    if (!/^\d*\s*\//.test(rest)) return [];
    const message = `Số và ký hiệu "${raw}" ghi theo cách của Chính quyền; chuẩn Đảng ghi "Số 05-HD/VPTW" (không dấu hai chấm, dấu "-" trước tên loại).`;
    return [paragraphIssue(p, { ruleId: "header.so.format", severity: "error", message, reference: source, actual: raw })];
  }
  const [, number, type, agency] = parts;
  const problems: string[] = [];
  if (head[1]) problems.push("không dùng dấu hai chấm sau \"Số\"");
  if (number.length === 1) problems.push("số nhỏ hơn 10 phải thêm số 0 phía trước");
  const expected = number === "" ? `Số ${rest}` : `Số ${pad2(number)}-${type.replace(/\s+/g, "")}/${agency.replace(/\s*-\s*/g, "-").replace(/\s+/g, "")}`;
  if (number !== "" && problems.length === 0 && expected !== raw) problems.push("không có khoảng trắng quanh dấu \"-\" và \"/\"");
  if (problems.length === 0) return [];
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

/** Các đoạn không trống trong bảng, gom theo ô. */
function tableCells(doc: CheckDocument): Map<string, CheckParagraph[]> {
  const cells = new Map<string, CheckParagraph[]>();
  for (const p of doc.paragraphs) {
    if (!p.table || p.resolved.text.trim() === "") continue;
    cells.set(cellKey(p), [...(cells.get(cellKey(p)) ?? []), p]);
  }
  return cells;
}

/** Ô khác cùng hàng với ô `key`. */
function siblingCell(cells: Map<string, CheckParagraph[]>, key: string): CheckParagraph[] | undefined {
  const [table, row, col] = key.split(":").map(Number);
  return [...cells.entries()].find(([k]) => {
    const [t, r, c] = k.split(":").map(Number);
    return t === table && r === row && c !== col;
  })?.[1];
}

/** Khối đầu: tiêu đề, cơ quan ban hành và dấu *, số và ký hiệu, ngày tháng, "Kính gửi". */
function checkHeader(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const issues: Issue[] = [];
  const title = profile.elements.tieuDe;
  const issuer = profile.elements.coQuanBanHanh;
  const kinhGui = profile.elements.kinhGuiTu;
  const date = profile.elements.diaDanhNgayThang;

  for (const p of doc.paragraphs) {
    if (!p.inTable || p.resolved.text.trim() === "") continue;
    const text = p.resolved.text.trim();
    issues.push(...checkTitleParagraph(p, profile.identifyingTitle, "header.tieuDe", "Tiêu đề", title.sizePt, title.source));
    if (matchTitle(text, TIEU_NGU) !== "no") {
      issues.push(paragraphIssue(p, { ruleId: "header.tieuNguChinhQuyen", severity: "error", message: "Văn bản của Đảng không dùng tiêu ngữ \"Độc lập - Tự do - Hạnh phúc\" của Chính quyền.", reference: HEADER_MARKS.source, actual: text }));
    }
    issues.push(...checkSo(p, profile.elements.soKyHieu.source));
    issues.push(...checkDate(p, date.source));
    if (DATE_PATTERN.test(text) && p.resolved.italic !== true) {
      issues.push(paragraphIssue(p, { ruleId: "header.date.italic", severity: "error", message: "Địa danh và ngày tháng phải in nghiêng.", reference: date.source, fix: fontFix(p, { italic: true }, "In nghiêng ngày tháng") }));
    }
    if (/^Kính (gửi|trình)\s*:$/.test(text) && p.resolved.italic !== true) {
      issues.push(paragraphIssue(p, { ruleId: "header.kinhGui.italic", severity: "error", message: `Từ "${text.replace(/:$/, "")}" phải in nghiêng.`, reference: kinhGui.source, fix: fontFix(p, { italic: true }, `In nghiêng "${text}"`) }));
    }
  }

  // Cơ quan ban hành: ô trái cùng hàng với tiêu đề; dòng ngay trước dấu * là tên cơ quan ban hành.
  const cells = tableCells(doc);
  for (const [key, cell] of cells) {
    if (!cell.some((p) => matchTitle(p.resolved.text, profile.identifyingTitle) !== "no")) continue;
    const left = siblingCell(cells, key);
    if (!left) continue;
    const star = left.findIndex((p) => p.resolved.text.trim() === HEADER_MARKS.agencyStar);
    if (star < 0) {
      issues.push(paragraphIssue(left[0], { ruleId: "header.star", severity: "warning", message: `Dưới tên cơ quan ban hành phải có dấu ${HEADER_MARKS.agencyStar}.`, reference: HEADER_MARKS.source }));
      continue;
    }
    if (star === 0) continue;
    const name = left[star - 1];
    const nameText = name.resolved.text.trim();
    if (nameText !== nameText.toUpperCase()) {
      issues.push(paragraphIssue(name, { ruleId: "header.coQuanBanHanh.upper", severity: "error", message: "Tên cơ quan ban hành phải in hoa.", reference: issuer.source, actual: nameText, expected: nameText.toUpperCase(), fix: replaceFix(name, nameText.toUpperCase(), "In hoa tên cơ quan ban hành") }));
    }
    if (name.resolved.bold !== true) {
      issues.push(paragraphIssue(name, { ruleId: "header.coQuanBanHanh.bold", severity: "error", message: "Tên cơ quan ban hành phải in đậm.", reference: issuer.source, fix: fontFix(name, { bold: true }, "In đậm tên cơ quan ban hành") }));
    }
    if (!inRange(name.resolved.sizePt, issuer.sizePt)) {
      issues.push(paragraphIssue(name, { ruleId: "header.coQuanBanHanh.size", severity: "error", message: `Tên cơ quan ban hành cỡ ${name.resolved.sizePt} pt; phải ${issuer.sizePt.min} pt.`, reference: issuer.source, fix: sizeFix(name, issuer.sizePt) }));
    }
  }
  return issues;
}

/** Thể loại, trích yếu và dòng 5 dấu gạch nối dưới trích yếu (văn bản có tên thể loại); công văn không có phần này. */
function checkTitleBlock(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const issues: Issue[] = [];
  const body = doc.paragraphs.filter((p) => !p.inTable && p.resolved.text.trim() !== "");
  const theLoai = profile.elements.theLoai;
  const trichYeu = profile.elements.trichYeu;
  const isCandidate = (p: CheckParagraph): boolean => {
    const text = p.resolved.text.trim();
    return p.resolved.align === "center" && text.length >= 3 && text === text.toUpperCase() && text !== text.toLowerCase() && !/^-+$/.test(text);
  };
  let at = body.slice(0, FIRST_TITLE_PARAGRAPHS).findIndex(isCandidate);
  if (at < 0) return issues;
  // Văn bản đại hội có dòng tên đại hội (in hoa, giữa) ngay trên thể loại: thể loại là dòng kế tiếp.
  while (CONGRESS_NAME.test(body[at].resolved.text.trim()) && body[at + 1] && isCandidate(body[at + 1])) at += 1;
  const type = body[at];
  if (!inRange(type.resolved.sizePt, theLoai.sizePt)) {
    issues.push(paragraphIssue(type, { ruleId: "title.theLoai.size", severity: "error", message: `Thể loại cỡ ${type.resolved.sizePt} pt; phải từ ${theLoai.sizePt.min} đến ${theLoai.sizePt.max} pt.`, reference: theLoai.source, fix: sizeFix(type, theLoai.sizePt) }));
  }
  if (type.resolved.bold !== true) {
    issues.push(paragraphIssue(type, { ruleId: "title.theLoai.bold", severity: "error", message: "Thể loại phải in đậm.", reference: theLoai.source, fix: fontFix(type, { bold: true }, "In đậm thể loại") }));
  }
  const subject = body[at + 1];
  if (!subject) return issues;
  if (subject.resolved.bold !== true) {
    issues.push(paragraphIssue(subject, { ruleId: "title.trichYeu.bold", severity: "error", message: "Trích yếu nội dung phải in đậm.", reference: trichYeu.source, fix: fontFix(subject, { bold: true }, "In đậm trích yếu") }));
  }
  if (!inRange(subject.resolved.sizePt, trichYeu.sizePt)) {
    issues.push(paragraphIssue(subject, { ruleId: "title.trichYeu.size", severity: "error", message: `Trích yếu cỡ ${subject.resolved.sizePt} pt; phải từ ${trichYeu.sizePt.min} đến ${trichYeu.sizePt.max} pt.`, reference: trichYeu.source, fix: sizeFix(subject, trichYeu.sizePt) }));
  }
  const dashes = body[at + 2];
  const wanted = "-".repeat(HEADER_MARKS.trichYeuDashCount);
  if (!dashes || !/^-+$/.test(dashes.resolved.text.trim())) {
    issues.push(paragraphIssue(subject, { ruleId: "title.dashes.missing", severity: "error", message: `Dưới trích yếu phải có dòng ${HEADER_MARKS.trichYeuDashCount} dấu gạch nối (${wanted}).`, reference: HEADER_MARKS.source }));
  } else if (dashes.resolved.text.trim() !== wanted) {
    issues.push(paragraphIssue(dashes, { ruleId: "title.dashes.count", severity: "error", message: `Dòng gạch nối dưới trích yếu phải đúng ${HEADER_MARKS.trichYeuDashCount} dấu.`, reference: HEADER_MARKS.source, actual: dashes.resolved.text.trim(), expected: wanted, fix: replaceFix(dashes, wanted, `Đặt ${HEADER_MARKS.trichYeuDashCount} dấu gạch nối`) }));
  }
  return issues;
}

/** Danh sách "Nơi nhận": gạch đầu dòng, ";" cuối mục, "." cuối mục cuối, cỡ chữ; dòng sau mục cuối (ký hiệu người soạn) bỏ qua. */
function checkRecipients(items: CheckParagraph[], profile: StandardProfile): Issue[] {
  const issues: Issue[] = [];
  const rule = profile.elements.noiNhanDanhSach;
  const lastDash = items.map((item) => item.resolved.text.trim().startsWith("-")).lastIndexOf(true);
  items.forEach((item, i) => {
    if (lastDash >= 0 && i > lastDash) return;
    const text = item.resolved.text.trim();
    const last = i === lastDash;
    const wanted = last ? "." : ";";
    if (!text.endsWith(wanted)) {
      issues.push(
        paragraphIssue(item, {
          ruleId: "footer.noiNhan.punctuation",
          severity: "error",
          message: last ? "Mục cuối của nơi nhận phải kết thúc bằng dấu chấm." : "Mỗi mục của nơi nhận phải kết thúc bằng dấu chấm phẩy.",
          reference: rule.source,
          actual: text.slice(-1),
          expected: wanted,
          fix: replaceFix(item, `${text.replace(/[;.,]+$/, "")}${wanted}`, last ? "Thêm dấu chấm cuối mục" : "Thêm dấu chấm phẩy cuối mục"),
        }),
      );
    }
    if (!text.startsWith("-")) {
      issues.push(paragraphIssue(item, { ruleId: "footer.noiNhan.dash", severity: "warning", message: "Mỗi mục của nơi nhận bắt đầu bằng dấu gạch ngang.", reference: rule.source }));
    }
    if (!inRange(item.resolved.sizePt, rule.sizePt)) {
      issues.push(paragraphIssue(item, { ruleId: "footer.noiNhan.size", severity: "error", message: `Danh sách nơi nhận cỡ ${item.resolved.sizePt} pt; phải ${rule.sizePt.min} pt.`, reference: rule.source, fix: sizeFix(item, rule.sizePt) }));
    }
  });
  return issues;
}

/** Ô chữ ký: quyền hạn, chức vụ (in hoa, không đậm), họ tên người ký. */
function checkSignature(lines: CheckParagraph[], profile: StandardProfile): Issue[] {
  const issues: Issue[] = [];
  const chucVu = profile.elements.chucVu;
  const quyenHan = profile.elements.quyenHan;
  const hoTen = profile.elements.hoTen;
  const name = lines[lines.length - 1];

  lines.slice(0, -1).forEach((line, i) => {
    const text = line.resolved.text.trim();
    if (text !== text.toUpperCase()) {
      issues.push(paragraphIssue(line, { ruleId: "footer.chucVu.upper", severity: "error", message: "Chức vụ của người ký phải in hoa.", reference: chucVu.source, actual: text, expected: text.toUpperCase(), fix: replaceFix(line, text.toUpperCase(), "In hoa chức vụ") }));
    }
    // Dòng đầu có quyền hạn thì quyền hạn đậm, chức vụ không đậm: không so đậm theo cả đoạn.
    const hasAuthority = i === 0 && AUTHORITY_ABBR.test(text);
    const rule = hasAuthority ? quyenHan : chucVu;
    if (!hasAuthority && line.resolved.bold === true) {
      issues.push(paragraphIssue(line, { ruleId: "footer.chucVu.bold", severity: "error", message: "Chức vụ của người ký không in đậm.", reference: chucVu.source, fix: fontFix(line, { bold: false }, "Bỏ in đậm chức vụ") }));
    }
    if (!inRange(line.resolved.sizePt, rule.sizePt)) {
      issues.push(paragraphIssue(line, { ruleId: "footer.chucVu.size", severity: "error", message: `Chức vụ cỡ ${line.resolved.sizePt} pt; phải ${rule.sizePt.min} pt.`, reference: rule.source, fix: sizeFix(line, rule.sizePt) }));
    }
    if (i !== 0) return;
    const gov = /^(TM|KT|TL)\.?(?=\s)/.exec(text);
    if (gov) {
      const wanted = AUTHORITY_GOV[gov[1]];
      issues.push(paragraphIssue(line, { ruleId: "footer.quyenHan.format", severity: "error", message: `Quyền hạn "${gov[0]}" phải viết "${wanted}" theo cách của Đảng.`, reference: quyenHan.source, actual: gov[0], expected: wanted, fix: replaceFix(line, `${wanted}${text.slice(gov[0].length)}`, "Sửa cách viết quyền hạn") }));
    } else if (/^TUQ\.?(?=\s)/.test(text)) {
      issues.push(paragraphIssue(line, { ruleId: "footer.quyenHan.format", severity: "error", message: "Chuẩn Đảng không có quyền hạn \"TUQ.\"; dùng T/M, K/T, T/L hoặc Q.", reference: quyenHan.source, actual: text.split(/\s/)[0] }));
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
  return issues;
}

/** Khối cuối: "Nơi nhận" gạch chân và danh sách, quyền hạn, chức vụ, họ tên người ký. */
function checkFooter(doc: CheckDocument, profile: StandardProfile): Issue[] {
  const issues: Issue[] = [];
  const cells = tableCells(doc);
  const noiNhanTu = profile.elements.noiNhanTu;

  for (const [key, cell] of cells) {
    const headingAt = cell.findIndex((p) => /^nơi nhận\s*:$/i.test(p.resolved.text.trim()));
    if (headingAt < 0) continue;
    const heading = cell[headingAt];
    const r = heading.resolved;
    if (r.underline !== true || r.bold === true || r.italic === true || !inRange(r.sizePt, noiNhanTu.sizePt)) {
      issues.push(
        paragraphIssue(heading, {
          ruleId: "footer.noiNhan.format",
          severity: "error",
          message: `Từ "Nơi nhận" phải gạch chân, không đậm, không nghiêng, cỡ ${noiNhanTu.sizePt.min} pt.`,
          reference: noiNhanTu.source,
          fix: fontFix(heading, { underline: true, bold: false, italic: false, sizePt: noiNhanTu.sizePt.min }, "Đặt định dạng cho \"Nơi nhận\""),
        }),
      );
    }
    issues.push(...checkRecipients(cell.slice(headingAt + 1), profile));

    const signCell = siblingCell(cells, key);
    if (!signCell) continue;
    const confirm = /^Xác nhận chữ ký/.test(signCell[signCell.length - 1].resolved.text.trim());
    const lines = confirm ? signCell.slice(0, -1) : signCell;
    if (lines.length >= 2) issues.push(...checkSignature(lines, profile));
  }
  return issues;
}

/** Quy tắc khối đầu, tiêu đề và khối cuối của chuẩn Đảng, dùng làm `extraRules` của engine. */
export function checkPartyHeaderFooter(doc: CheckDocument, profile: StandardProfile): Issue[] {
  return [...checkHeader(doc, profile), ...checkTitleBlock(doc, profile), ...checkFooter(doc, profile)];
}
