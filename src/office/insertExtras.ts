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

/** Đếm phụ lục theo nhãn tự đánh số "Phụ lục ..." của các đoạn danh sách không có chữ (phụ lục dựng trực tiếp, không style). */
export async function countAppendicesByLabel(): Promise<number> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/isListItem,items/text");
    await context.sync();
    const candidates = paragraphs.items.filter((paragraph) => paragraph.isListItem && paragraph.text.trim() === "");
    candidates.forEach((paragraph) => paragraph.listItem.load("listString"));
    await context.sync();
    return candidates.filter((paragraph) => paragraph.listItem.listString.startsWith("Phụ lục")).length;
  });
}

export interface AppendOptions {
  /** Đếm các phụ lục đã có. */
  count: () => Promise<number>;
  /** Đoạn nào chứng tỏ phụ lục mới đã vào tài liệu (so tên phụ lục). */
  matches: (paragraph: { style: string; text: string }) => boolean;
  /** Dựng XML thân và phần numbering cho phụ lục có số thứ tự `number`. */
  build: (number: number) => { bodyXml: string; numberingXml: string };
}

/** Chèn phụ lục vào cuối tài liệu và đọc lại để kiểm tra. */
export async function appendAppendix(options: AppendOptions): Promise<ExtrasResult & { number: number }> {
  const number = (await options.count()) + 1;
  const { bodyXml, numberingXml } = options.build(number);
  return Word.run(async (context) => {
    context.document.body.insertOoxml(flatOpc({ bodyXml, numberingXml }), "End");
    await context.sync();
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/style,items/text");
    await context.sync();
    const ok = paragraphs.items.some((paragraph) => options.matches({ style: paragraph.style, text: paragraph.text }));
    return { number, verified: ok, note: ok ? `Đã chèn phụ lục số ${number} ở cuối tài liệu.` : "Đã chèn nhưng đọc lại không thấy tên phụ lục; hãy kiểm tra tài liệu." };
  });
}

/** Chèn đoạn chữ tại con trỏ rồi đọc lại thân văn bản xem có chữ mong đợi không. */
export async function insertTextAtSelection(bodyXml: string, expectedText: string): Promise<ExtrasResult> {
  return Word.run(async (context) => {
    context.document.getSelection().insertOoxml(flatOpc({ bodyXml }), "Replace");
    await context.sync();
    context.document.body.load("text");
    await context.sync();
    const ok = context.document.body.text.includes(expectedText);
    return { verified: ok, note: ok ? "Đã chèn tại con trỏ." : "Đã chèn nhưng đọc lại không thấy chữ; hãy kiểm tra tài liệu." };
  });
}

/** Chèn vào cuối tài liệu rồi đọc lại xem có chữ mong đợi không. */
export async function appendToEnd(bodyXml: string, expectedText: string): Promise<ExtrasResult> {
  return Word.run(async (context) => {
    context.document.body.insertOoxml(flatOpc({ bodyXml }), "End");
    await context.sync();
    context.document.body.load("text");
    await context.sync();
    const ok = context.document.body.text.includes(expectedText);
    return { verified: ok, note: ok ? "Đã chèn ở cuối tài liệu." : "Đã chèn nhưng đọc lại không thấy chữ; hãy kiểm tra tài liệu." };
  });
}
