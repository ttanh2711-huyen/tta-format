import { roleStyle, type ParagraphRole } from "../office/paragraphRoles";
import { buildStyleSpecs } from "../ooxml/styleSpecs";
import { STANDARDS } from "../standards";
import type { StandardProfile } from "../standards/types";

/**
 * Mô-đun cấu hình của bộ macro VBA (T29), sinh từ chính hồ sơ chuẩn mà add-in dùng (khổ giấy, lề, cỡ chữ, style,
 * giãn dòng, style theo vai trò đoạn) để hai bản không lệch nhau. Chuẩn đánh số theo thứ tự `STANDARDS` (1, 2…).
 */

const ROLES: readonly ParagraphRole[] = ["chuong", "tenChuong", "muc", "tenMuc", "dieu", "deMuc", "khoan", "diem", "y", "noiDung"];

/** Cỡ chữ các thành phần khối đầu, khối cuối theo khoá chung cho cả hai chuẩn. */
const SIZE_KEYS: Record<string, Record<StandardProfile["headerKind"], string | null>> = {
  title: { quocHieu: "quocHieu", dang: "tieuDe" },
  motto: { quocHieu: "tieuNgu", dang: null },
  parent: { quocHieu: "tenCoQuanChuQuan", dang: "coQuanCapTren" },
  issuer: { quocHieu: "tenCoQuanBanHanh", dang: "coQuanBanHanh" },
  so: { quocHieu: "so", dang: "soKyHieu" },
  date: { quocHieu: "diaDanhNgayThang", dang: "diaDanhNgayThang" },
  vv: { quocHieu: "trichYeuCongVan", dang: "trichYeuCongVan" },
  docType: { quocHieu: "tenLoai", dang: "theLoai" },
  subject: { quocHieu: "trichYeu", dang: "trichYeu" },
  authority: { quocHieu: "quyenHanChucVu", dang: "quyenHan" },
  signTitle: { quocHieu: "quyenHanChucVu", dang: "chucVu" },
  signName: { quocHieu: "hoTenNguoiKy", dang: "hoTen" },
  recipientsLabel: { quocHieu: "noiNhanTieuDe", dang: "noiNhanTu" },
  recipientsList: { quocHieu: "noiNhanDanhSach", dang: "noiNhanDanhSach" },
};

const ALIGN: Record<string, number> = { Left: 0, Centered: 1, Right: 2, Justified: 3 };
const q = (text: string): string => `"${text.replace(/"/g, '""')}"`;
const num = (n: number): string => String(Math.round(n * 100) / 100);
const bool = (b: boolean): string => (b ? "True" : "False");

function selectFunction(name: string, signature: string, returns: string, cases: Array<[string, string]>, fallback: string): string {
  const body = cases.map(([when, value]) => `        Case ${when}: ${name} = ${value}`).join("\n");
  return `Public Function ${name}(${signature}) As ${returns}\n    Select Case std\n${body}\n        Case Else: ${name} = ${fallback}\n    End Select\nEnd Function\n`;
}

export function configModule(): string {
  const stds = STANDARDS.map((profile, i) => ({ profile, n: i + 1, specs: buildStyleSpecs(profile) }));
  const out: string[] = [
    'Attribute VB_Name = "TTA_Config"',
    "' TTA-Format cho Word 2007 trở về sau: cấu hình sinh tự động từ hồ sơ chuẩn của add-in. KHÔNG sửa tay;",
    "' chạy lại: UPDATE_VBA=1 npx jest src/vba",
    "Option Explicit",
    "",
    `Public Const TTA_STANDARD_COUNT As Integer = ${stds.length}`,
    "",
  ];
  out.push(selectFunction("CfgName", "ByVal std As Integer", "String", stds.map(({ n, profile }) => [String(n), q(profile.name)]), '""'));
  out.push(selectFunction("CfgPrefix", "ByVal std As Integer", "String", stds.map(({ n, profile }) => [String(n), q(profile.styles.prefix)]), '""'));
  out.push(selectFunction("CfgKind", "ByVal std As Integer", "Integer", stds.map(({ n, profile }) => [String(n), profile.headerKind === "quocHieu" ? "1" : "2"]), "1"));
  out.push(selectFunction("CfgTitle", "ByVal std As Integer", "String", stds.map(({ n, profile }) => [String(n), q(profile.identifyingTitle)]), '""'));
  out.push(selectFunction("CfgMotto", "ByVal std As Integer", "String", stds.map(({ n, profile }) => [String(n), q(profile.headerKind === "quocHieu" ? "Độc lập - Tự do - Hạnh phúc" : "")]), '""'));
  out.push(selectFunction("CfgFont", "ByVal std As Integer", "String", stds.map(({ n, profile }) => [String(n), q(profile.font.name)]), '"Times New Roman"'));
  out.push(selectFunction("CfgHideFirst", "ByVal std As Integer", "Boolean", stds.map(({ n, profile }) => [String(n), bool(profile.pageNumber.hideOnFirstPage)]), "True"));
  out.push(selectFunction("CfgPageNumberSize", "ByVal std As Integer", "Single", stds.map(({ n, profile }) => [String(n), num(profile.pageNumber.sizePt.max)]), "13"));

  // Trang: giá trị mặc định để đặt, khoảng cho phép để kiểm tra.
  out.push(
    "Public Sub CfgPage(ByVal std As Integer, ByRef topMm As Double, ByRef bottomMm As Double, ByRef leftMm As Double, ByRef rightMm As Double, ByRef headerMm As Double, ByRef footerMm As Double)\n    Select Case std\n" +
      stds
        .map(({ n, profile }) => {
          const d = profile.pageDefaults;
          return `        Case ${n}: topMm = ${num(d.marginsMm.top)}: bottomMm = ${num(d.marginsMm.bottom)}: leftMm = ${num(d.marginsMm.left)}: rightMm = ${num(d.marginsMm.right)}: headerMm = ${num(d.headerDistanceMm)}: footerMm = ${num(d.footerDistanceMm)}`;
        })
        .join("\n") +
      "\n    End Select\nEnd Sub\n",
  );
  out.push(
    "Public Sub CfgMarginRange(ByVal std As Integer, ByVal side As String, ByRef minMm As Double, ByRef maxMm As Double)\n    Select Case std & \"|\" & side\n" +
      stds
        .flatMap(({ n, profile }) => (["top", "bottom", "left", "right"] as const).map((side) => `        Case "${n}|${side}": minMm = ${num(profile.page.marginsMm[side].min)}: maxMm = ${num(profile.page.marginsMm[side].max)}`))
        .join("\n") +
      "\n    End Select\nEnd Sub\n",
  );

  // Nội dung: cỡ chữ, thụt dòng đầu, giãn dòng để kiểm tra.
  out.push(
    "Public Sub CfgBody(ByVal std As Integer, ByRef sizeMin As Single, ByRef sizeMax As Single, ByRef indentMin As Double, ByRef indentMax As Double, ByRef lineKind As Integer, ByRef lineMin As Double, ByRef lineMax As Double)\n    Select Case std\n" +
      stds
        .map(({ n, profile }) => {
          const size = profile.elements.noiDung.sizePt;
          const c = profile.content;
          const line = c.lineSpacing.kind === "multiple" ? `lineKind = 1: lineMin = ${num(c.lineSpacing.min)}: lineMax = ${num(c.lineSpacing.max)}` : `lineKind = 2: lineMin = ${num(c.lineSpacing.minPt)}: lineMax = ${num(c.lineSpacing.maxPt)}`;
          return `        Case ${n}: sizeMin = ${num(size.min)}: sizeMax = ${num(size.max)}: indentMin = ${num(Math.min(...c.firstLineIndentMm) - c.firstLineIndentToleranceMm)}: indentMax = ${num(Math.max(...c.firstLineIndentMm) + c.firstLineIndentToleranceMm)}: ${line}`;
        })
        .join("\n") +
      "\n    End Select\nEnd Sub\n",
  );

  // Style của chuẩn.
  out.push(selectFunction("CfgStyleCount", "ByVal std As Integer", "Integer", stds.map(({ n, specs }) => [String(n), String(specs.length)]), "0"));
  out.push(
    "Public Sub CfgStyle(ByVal std As Integer, ByVal i As Integer, ByRef styleName As String, ByRef sizePt As Single, ByRef isBold As Boolean, ByRef isItalic As Boolean, ByRef isUnderline As Boolean, ByRef align As Integer, ByRef firstIndentPt As Single, ByRef beforePt As Single, ByRef afterPt As Single, ByRef lineKind As Integer, ByRef lineValue As Single, ByRef keepNext As Boolean)\n    Select Case std * 100 + i\n" +
      stds
        .flatMap(({ n, specs }) =>
          specs.map((s, i) => {
            const line = s.paragraph.lineSpacing.kind === "multiple" ? `lineKind = 1: lineValue = ${num(s.paragraph.lineSpacing.lines)}` : `lineKind = 2: lineValue = ${num(s.paragraph.lineSpacing.pt)}`;
            return `        Case ${n * 100 + i + 1}: styleName = ${q(s.name)}: sizePt = ${num(s.font.sizePt)}: isBold = ${bool(s.font.bold)}: isItalic = ${bool(s.font.italic)}: isUnderline = ${bool(s.font.underline)}: align = ${ALIGN[s.paragraph.alignment] ?? 0}: firstIndentPt = ${num(s.paragraph.firstLineIndentPt)}: beforePt = ${num(s.paragraph.spaceBeforePt)}: afterPt = ${num(s.paragraph.spaceAfterPt)}: ${line}: keepNext = ${bool(s.paragraph.keepWithNext)}`;
          }),
        )
        .join("\n") +
      "\n    End Select\nEnd Sub\n",
  );

  // Style theo vai trò đoạn (cùng bảng với add-in, issue #3).
  out.push(
    "Public Function CfgRoleStyle(ByVal std As Integer, ByVal role As String) As String\n    Select Case std & \"|\" & role\n" +
      stds.flatMap(({ n, profile }) => ROLES.map((role) => `        Case "${n}|${role}": CfgRoleStyle = ${q(roleStyle(role, profile).styleName)}`)).join("\n") +
      `\n        Case Else: CfgRoleStyle = CfgPrefix(std) & "NoiDung"\n    End Select\nEnd Function\n`,
  );
  out.push(
    "Public Function CfgRoleBold(ByVal std As Integer, ByVal role As String) As Boolean\n    Select Case std & \"|\" & role\n" +
      stds.flatMap(({ n, profile }) => ROLES.filter((role) => roleStyle(role, profile).bold).map((role) => `        Case "${n}|${role}": CfgRoleBold = True`)).join("\n") +
      "\n        Case Else: CfgRoleBold = False\n    End Select\nEnd Function\n",
  );

  // Cỡ chữ khối đầu, khối cuối.
  out.push(
    "Public Function CfgSize(ByVal std As Integer, ByVal key As String) As Single\n    Select Case std & \"|\" & key\n" +
      stds
        .flatMap(({ n, profile }) =>
          Object.entries(SIZE_KEYS).map(([key, ids]) => {
            const id = ids[profile.headerKind];
            const size = id === null ? 0 : profile.styles.blockSizes.find((b) => b.elementId === id)?.defaultPt ?? 13;
            return `        Case "${n}|${key}": CfgSize = ${num(size)}`;
          }),
        )
        .join("\n") +
      "\n        Case Else: CfgSize = 13\n    End Select\nEnd Function\n",
  );
  return out.join("\n");
}
