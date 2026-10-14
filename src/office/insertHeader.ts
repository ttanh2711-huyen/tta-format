import { flatOpc } from "../ooxml/flatOpc";
import { readDocDefaults } from "./docDefaults";

/** Chèn khối đầu văn bản cần WordApi 1.3 (đọc bảng sau khi chèn để kiểm tra). */
export function isInsertHeaderSupported(): boolean {
  return Office.context.requirements.isSetSupported("WordApi", "1.3");
}

/** Tài liệu đang bắt đầu bằng một bảng (nhiều khả năng đã có khối đầu). */
export async function startsWithTable(): Promise<boolean> {
  return Word.run(async (context) => {
    const first = context.document.body.paragraphs.getFirst();
    first.load("tableNestingLevel");
    await context.sync();
    return first.tableNestingLevel > 0;
  });
}

export interface InsertResult {
  /** Bảng đầu tiên sau khi chèn có chứa chữ mong đợi. */
  verified: boolean;
  tables: number;
}

/**
 * Chèn bảng khối đầu vào đầu thân văn bản, kèm một đoạn trống phía sau (Word luôn cần đoạn sau bảng). Đọc lại
 * bảng đầu tiên để kiểm tra có đúng chữ `expectedText`. Một lần Word.run nên Ctrl+Z hoàn tác được.
 */
export async function insertHeaderBlock(bodyXml: string, expectedText: string): Promise<InsertResult> {
  return Word.run(async (context) => {
    const docDefaultsXml = await readDocDefaults(context);
    context.document.body.insertOoxml(flatOpc({ bodyXml: `${bodyXml}<w:p/>`, docDefaultsXml }), "Start");
    await context.sync();
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    const first = tables.items[0];
    if (!first) return { verified: false, tables: 0 };
    first.load("values");
    await context.sync();
    return { verified: first.values.some((row) => row.some((text) => text.includes(expectedText))), tables: tables.items.length };
  });
}
