import { flatOpc } from "../ooxml/flatOpc";

/** Chèn dấu khẩn và phụ lục cần WordApi 1.3 (đọc lại kết quả để kiểm tra). */
export function isInsertExtrasSupported(): boolean {
  return Office.context.requirements.isSetSupported("WordApi", "1.3");
}

export interface ExtrasResult {
  verified: boolean;
  note: string;
}

/** Chèn khung dấu tại vị trí con trỏ (thay vùng đang chọn). Một lần Word.run nên Ctrl+Z hoàn tác được. */
export async function insertAtSelection(bodyXml: string, expectedText: string): Promise<ExtrasResult> {
  return Word.run(async (context) => {
    const before = context.document.body.tables;
    before.load("items");
    await context.sync();
    const count = before.items.length;
    context.document.getSelection().insertOoxml(flatOpc({ bodyXml }), "Replace");
    await context.sync();
    const tables = context.document.body.tables;
    tables.load("items/values");
    await context.sync();
    const found = tables.items.some((table) => table.values.some((row) => row.some((text) => text.includes(expectedText))));
    return tables.items.length > count && found ? { verified: true, note: "Đã chèn khung tại con trỏ." } : { verified: false, note: "Đã chèn nhưng đọc lại không thấy khung; hãy kiểm tra tài liệu." };
  });
}

/** Số phụ lục đã có, đếm theo số đoạn mang style nhãn phụ lục của chuẩn. */
export async function countAppendices(labelStyle: string): Promise<number> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/style");
    await context.sync();
    return paragraphs.items.filter((paragraph) => paragraph.style === labelStyle).length;
  });
}

/** Chèn phụ lục vào cuối tài liệu; `build` nhận số thứ tự của phụ lục mới và trả XML thân cùng phần numbering. */
export async function appendAppendix(labelStyle: string, titleStyle: string, expectedTitle: string, build: (number: number) => { bodyXml: string; numberingXml: string }): Promise<ExtrasResult & { number: number }> {
  const number = (await countAppendices(labelStyle)) + 1;
  const { bodyXml, numberingXml } = build(number);
  return Word.run(async (context) => {
    context.document.body.insertOoxml(flatOpc({ bodyXml, numberingXml }), "End");
    await context.sync();
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/style,items/text");
    await context.sync();
    const ok = paragraphs.items.some((paragraph) => paragraph.style === titleStyle && paragraph.text.trim() === expectedTitle);
    return { number, verified: ok, note: ok ? `Đã chèn phụ lục số ${number} ở cuối tài liệu.` : "Đã chèn nhưng đọc lại không thấy tên phụ lục; hãy kiểm tra tài liệu." };
  });
}
