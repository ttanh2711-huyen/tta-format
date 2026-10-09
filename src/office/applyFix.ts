import type { Fix } from "../checker/types";

export interface ApplyFixResult {
  applied: number;
  /** Cách sửa không áp dụng được, kèm lý do. */
  skipped: string[];
}

/**
 * Áp dụng các cách sửa lên tài liệu trong một lần Word.run (nên Ctrl+Z hoàn tác được trọn một lần). Mỗi cách sửa chỉ
 * đụng đoạn có chỉ số `index`: thay chữ giữ định dạng của chữ được thay, hoặc đặt lại đậm, nghiêng, cỡ chữ của đoạn.
 */
export async function applyFixes(fixes: readonly Fix[]): Promise<ApplyFixResult> {
  if (fixes.length === 0) return { applied: 0, skipped: [] };
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/text");
    await context.sync();
    const skipped: string[] = [];
    let applied = 0;

    const searches = fixes.map((fix) => {
      const paragraph = paragraphs.items[fix.index];
      if (!paragraph) return undefined;
      if (fix.kind === "font") return { fix, paragraph, found: undefined };
      const found = paragraph.search(fix.find, { matchCase: true });
      found.load("items");
      return { fix, paragraph, found };
    });
    await context.sync();

    searches.forEach((entry, i) => {
      const fix = fixes[i];
      if (!entry) {
        skipped.push(`${fix.label}: không còn đoạn số ${fix.index + 1}`);
        return;
      }
      if (entry.fix.kind === "font") {
        const font = entry.paragraph.font;
        if (entry.fix.bold !== undefined) font.bold = entry.fix.bold;
        if (entry.fix.italic !== undefined) font.italic = entry.fix.italic;
        if (entry.fix.sizePt !== undefined) font.size = entry.fix.sizePt;
        applied += 1;
      } else if (entry.found && entry.found.items.length > 0) {
        entry.found.items[0].insertText(entry.fix.replacement, "Replace");
        applied += 1;
      } else {
        skipped.push(`${fix.label}: không tìm thấy chữ cần sửa (có thể đã được sửa)`);
      }
    });
    await context.sync();
    return { applied, skipped };
  });
}
