import type { Fix } from "../checker/types";
import { mmToPt } from "../standards/units";
import { rewriteWithExact } from "./formatBody";

export interface ApplyFixResult {
  applied: number;
  /** Cách sửa không áp dụng được, kèm lý do. */
  skipped: string[];
}

/**
 * Áp dụng các cách sửa lên tài liệu trong một lần Word.run (nên Ctrl+Z hoàn tác được trọn một lần). Cách sửa trên
 * đoạn chỉ đụng đoạn có chỉ số `index`; cách sửa trang chỉ đụng phần có số thứ tự `section`. Giãn dòng Exactly phải
 * ghi lại đoạn bằng OOXML (API không đặt được), nên bình luận gắn trong đoạn đó có thể không còn.
 */
export async function applyFixes(fixes: readonly Fix[]): Promise<ApplyFixResult> {
  if (fixes.length === 0) return { applied: 0, skipped: [] };
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/text");
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    const skipped: string[] = [];
    let applied = 0;

    const searches = fixes.map((fix) => {
      if (fix.kind !== "replace") return undefined;
      const paragraph = paragraphs.items[fix.index];
      if (!paragraph) return undefined;
      const found = paragraph.search(fix.find, { matchCase: true });
      found.load("items");
      return found;
    });
    await context.sync();

    const exact: Array<{ paragraph: Word.Paragraph; pt: number }> = [];
    fixes.forEach((fix, i) => {
      if (fix.kind === "page") {
        const section = sections.items[fix.section - 1];
        if (!section) {
          skipped.push(`${fix.label}: không còn phần ${fix.section}`);
          return;
        }
        if (fix.paper) section.pageSetup.set({ paperSize: fix.paper, orientation: "Portrait" });
        applied += 1;
        return;
      }
      const paragraph = paragraphs.items[fix.index];
      if (!paragraph) {
        skipped.push(`${fix.label}: không còn đoạn số ${fix.index + 1}`);
        return;
      }
      if (fix.kind === "font") {
        if (fix.name !== undefined) paragraph.font.name = fix.name;
        if (fix.bold !== undefined) paragraph.font.bold = fix.bold;
        if (fix.italic !== undefined) paragraph.font.italic = fix.italic;
        if (fix.underline !== undefined) paragraph.font.underline = fix.underline ? "Single" : "None";
        if (fix.sizePt !== undefined) paragraph.font.size = fix.sizePt;
        applied += 1;
      } else if (fix.kind === "paragraph") {
        if (fix.align) paragraph.alignment = "Justified";
        if (fix.firstLineIndentMm !== undefined) paragraph.firstLineIndent = mmToPt(fix.firstLineIndentMm);
        if (fix.spaceBeforePt !== undefined) paragraph.spaceBefore = fix.spaceBeforePt;
        if (fix.spaceAfterPt !== undefined) paragraph.spaceAfter = fix.spaceAfterPt;
        if (fix.lineMultiple !== undefined) paragraph.lineSpacing = fix.lineMultiple * 12;
        applied += 1;
      } else if (fix.kind === "exact") {
        exact.push({ paragraph, pt: fix.pt });
        applied += 1;
      } else {
        const found = searches[i];
        if (found && found.items.length > 0) {
          found.items[0].insertText(fix.replacement, "Replace");
          applied += 1;
        } else {
          skipped.push(`${fix.label}: không tìm thấy chữ cần sửa (có thể đã được sửa)`);
        }
      }
    });
    await context.sync();

    // Lề đặt sau khổ giấy: đổi khổ giấy có thể làm Word tính lại lề.
    for (const fix of fixes) {
      if (fix.kind !== "page") continue;
      const section = sections.items[fix.section - 1];
      if (!section) continue;
      const m = fix.marginsMm ?? {};
      section.pageSetup.set({
        ...(m.top === undefined ? {} : { topMargin: mmToPt(m.top) }),
        ...(m.bottom === undefined ? {} : { bottomMargin: mmToPt(m.bottom) }),
        ...(m.left === undefined ? {} : { leftMargin: mmToPt(m.left) }),
        ...(m.right === undefined ? {} : { rightMargin: mmToPt(m.right) }),
        ...(fix.headerDistanceMm === undefined ? {} : { headerDistance: mmToPt(fix.headerDistanceMm) }),
      });
    }
    await context.sync();

    for (const { paragraph, pt } of exact) await rewriteWithExact(context, [paragraph], pt);
    return { applied, skipped };
  });
}
