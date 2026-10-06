export const SAMPLE_TEXT =
  "Thử nghiệm NĐ 30: CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM - Độc lập - Tự do - Hạnh phúc";

/** Phần của Word.Body mà hàm cần, để kiểm thử được mà không cần Word. */
export interface ParagraphInserter {
  insertParagraph(text: string, insertLocation: "End"): unknown;
}

export function insertSampleParagraph(body: ParagraphInserter): unknown {
  return body.insertParagraph(SAMPLE_TEXT, "End");
}
