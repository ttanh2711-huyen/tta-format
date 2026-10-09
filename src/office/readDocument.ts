import { buildCheckDocument } from "../checker/documentModel";
import type { CheckDocument, PageInfo } from "../checker/types";
import { parseStyleTable } from "../ooxml/styleResolver";
import { extractDocModel } from "../rules/docModel";

const PT_PER_MM = 72 / 25.4;
const PAGE_PROPS = "pageWidth,pageHeight,topMargin,bottomMargin,leftMargin,rightMargin,headerDistance,footerDistance,mirrorMargins";

const mm = (pt: number): number => Math.round((pt / PT_PER_MM) * 100) / 100;

/** Đọc tài liệu cho bộ kiểm tra cần WordApiDesktop 1.3 (thiết lập trang từng phần). */
export function isReadDocumentSupported(): boolean {
  return Office.context.requirements.isSetSupported("WordApiDesktop", "1.3");
}

/**
 * Đọc tài liệu đang mở: OOXML của thân (đoạn, bảng, style) và thiết lập trang của từng phần qua API. Chỉ đọc,
 * không sửa gì. Chỉ số đoạn khớp `body.paragraphs` để nhảy tới vị trí lỗi.
 */
export async function readCheckDocument(): Promise<CheckDocument> {
  return Word.run(async (context) => {
    const ooxml = context.document.body.getOoxml();
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    for (const section of sections.items) section.pageSetup.load(PAGE_PROPS);
    await context.sync();

    const pages: PageInfo[] = sections.items.map((section) => {
      const ps = section.pageSetup;
      return {
        widthMm: mm(ps.pageWidth),
        heightMm: mm(ps.pageHeight),
        marginsMm: { top: mm(ps.topMargin), bottom: mm(ps.bottomMargin), left: mm(ps.leftMargin), right: mm(ps.rightMargin) },
        headerDistanceMm: mm(ps.headerDistance),
        footerDistanceMm: mm(ps.footerDistance),
        mirrorMargins: ps.mirrorMargins,
      };
    });
    const stylesXml = /<w:styles\b[\s\S]*<\/w:styles>/.exec(ooxml.value)?.[0] ?? "";
    return buildCheckDocument(extractDocModel({ documentXml: ooxml.value }), parseStyleTable(stylesXml), pages);
  });
}

/** Chọn đoạn thứ `index` của thân văn bản để người dùng thấy ngay chỗ lỗi. */
export async function selectParagraph(index: number): Promise<boolean> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();
    const target = paragraphs.items[index];
    if (!target) return false;
    target.select();
    await context.sync();
    return true;
  });
}
