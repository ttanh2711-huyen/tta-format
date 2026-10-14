import { flatOpc } from "../ooxml/flatOpc";
import { readDocDefaults } from "./docDefaults";
import type { InsertResult } from "./insertHeader";

/**
 * Chèn khối cuối (bảng nơi nhận và chữ ký) vào cuối thân văn bản, rồi đọc lại bảng cuối cùng để kiểm tra
 * có chứa chữ `expectedText`. Một lần Word.run nên Ctrl+Z hoàn tác được.
 */
export async function insertFooterBlock(bodyXml: string, expectedText: string): Promise<InsertResult> {
  return Word.run(async (context) => {
    const docDefaultsXml = await readDocDefaults(context);
    context.document.body.insertOoxml(flatOpc({ bodyXml, docDefaultsXml }), "End");
    await context.sync();
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    const last = tables.items[tables.items.length - 1];
    if (!last) return { verified: false, tables: 0 };
    last.load("values");
    await context.sync();
    return { verified: last.values.some((row) => row.some((text) => text.includes(expectedText))), tables: tables.items.length };
  });
}
