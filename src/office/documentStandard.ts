import type { StandardId } from "../standards/types";

const STANDARD_KEY = "nd30.standard";
const NO_ASK_KEY = "nd30.standardNoAsk";

/** Những gì cần biết về tài liệu đang mở để quyết định chuẩn. */
export interface DocumentSnapshot {
  /** Chuẩn đã ghi trong thuộc tính tuỳ chỉnh (chưa kiểm tra hợp lệ). */
  stored: string | null;
  noAsk: boolean;
  /** Chữ các đoạn từ đầu tài liệu (kể cả đoạn trong bảng), dùng để tìm tiêu đề. */
  headerTexts: string[];
  /** Có chữ hoặc bảng trong tài liệu không. */
  hasContent: boolean;
}

/** Thuộc tính tuỳ chỉnh của tài liệu cần WordApi 1.3. */
export function isDocumentStandardSupported(): boolean {
  return Office.context.requirements.isSetSupported("WordApi", "1.3");
}

/** Số đoạn đầu tài liệu được đọc; đủ rộng để chứa bảng đầu văn bản và vài dòng sau đó. */
const HEADER_READ_LIMIT = 80;

export async function readDocumentSnapshot(): Promise<DocumentSnapshot> {
  return Word.run(async (context) => {
    const standard = context.document.properties.customProperties.getItemOrNullObject(STANDARD_KEY);
    const noAsk = context.document.properties.customProperties.getItemOrNullObject(NO_ASK_KEY);
    standard.load("isNullObject,value");
    noAsk.load("isNullObject,value");
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/text");
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    const texts = paragraphs.items.map((paragraph) => paragraph.text);
    return {
      stored: standard.isNullObject ? null : String(standard.value),
      noAsk: !noAsk.isNullObject && String(noAsk.value) === "1",
      headerTexts: texts.slice(0, HEADER_READ_LIMIT),
      hasContent: tables.items.length > 0 || texts.some((text) => text.trim() !== ""),
    };
  });
}

async function writeProperty(key: string, value: string): Promise<void> {
  await Word.run(async (context) => {
    context.document.properties.customProperties.add(key, value);
    await context.sync();
  });
}

/** Ghi chuẩn vào tài liệu để lần mở sau không phải nhận diện lại. */
export async function writeDocumentStandard(id: StandardId): Promise<void> {
  await writeProperty(STANDARD_KEY, id);
}

export async function writeDocumentNoAsk(): Promise<void> {
  await writeProperty(NO_ASK_KEY, "1");
}
