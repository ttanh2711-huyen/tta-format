/** Kiểm tra kết quả của lượt đặt giãn dòng Exactly trực tiếp trên đoạn (xem exactSpacing.ts). */

/** Số đoạn trong phần thân có giãn dòng Exactly đúng `twips`. */
export function countExactParagraphs(ooxml: string, twips: number): number {
  const body = /<w:body>([\s\S]*)<\/w:body>/.exec(ooxml);
  const paragraphs = (body ? body[1] : "").match(/<w:p\b[\s\S]*?(?:<\/w:p>|\/>)/g) ?? [];
  const spacing = new RegExp(String.raw`<w:spacing\b(?=[^>]*w:lineRule="exact")(?=[^>]*w:line="${twips}")[^>]*>`);
  return paragraphs.filter((paragraph) => spacing.test(paragraph)).length;
}

export type ExactVerdict = { ok: true } | { ok: false; problem: string };

export interface ExactVerifyInput {
  /** Chữ từng đoạn trước và sau khi ghi lại. */
  before: readonly string[];
  after: readonly string[];
  /** OOXML của vùng đã xử lý, đọc lại sau khi ghi. */
  ooxml: string;
  twips: number;
  expectedExact: number;
}

export function verifyExactPass(input: ExactVerifyInput): ExactVerdict {
  const { before, after } = input;
  if (before.length !== after.length) {
    return { ok: false, problem: `số đoạn đổi từ ${before.length} thành ${after.length}` };
  }
  const changed = before.findIndex((text, index) => text !== after[index]);
  if (changed >= 0) {
    return { ok: false, problem: `nội dung đoạn ${changed + 1} bị thay đổi` };
  }
  const found = countExactParagraphs(input.ooxml, input.twips);
  if (found < input.expectedExact) {
    return { ok: false, problem: `chỉ ${found} trên ${input.expectedExact} đoạn có giãn dòng Exactly` };
  }
  return { ok: true };
}
