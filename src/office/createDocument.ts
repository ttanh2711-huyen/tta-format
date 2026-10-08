import { assembleDocumentBody, BODY_MARKER, type AssembleInput } from "../ooxml/documentAssembler";
import { flatOpc } from "../ooxml/flatOpc";
import { QUOC_HIEU } from "../ooxml/headerBlock";
import { pageNumberStep, pageStep, styleStep, type FixAllContext } from "../pipeline/steps";
import { runPipeline, type PipelineReport, type PipelineStep, type ProgressEvent } from "../pipeline/runner";

/** Ngữ cảnh của trình tạo công văn: các tuỳ chọn của "Chỉnh toàn bộ" cộng với nội dung đã nhập. */
export interface LetterContext extends FixAllContext {
  document: AssembleInput;
}

const insertStep: PipelineStep<LetterContext> = {
  id: "insert",
  label: "3. Dựng khối đầu, Kính gửi, khối cuối",
  order: 30,
  async plan() {
    return { willRun: true, lines: ["Thay toàn bộ nội dung bằng công văn mới: khối đầu, Kính gửi, chỗ nội dung, khối cuối."] };
  },
  async run(ctx) {
    const bodyStyle = `${ctx.profile.styles.prefix}NoiDung`;
    return Word.run(async (context) => {
      const body = context.document.body;
      body.insertOoxml(flatOpc({ bodyXml: assembleDocumentBody(ctx.document) }), "Replace");
      await context.sync();

      const paragraphs = body.paragraphs;
      paragraphs.load("items/text");
      const tables = body.tables;
      tables.load("items");
      await context.sync();
      const marker = paragraphs.items.find((paragraph) => paragraph.text.includes(BODY_MARKER));
      if (!marker) throw new Error("Không thấy chỗ đánh dấu nội dung sau khi chèn.");
      marker.insertText("", "Replace");
      marker.style = bodyStyle;
      marker.select();
      const first = tables.items[0];
      const last = tables.items[tables.items.length - 1];
      first.load("values");
      last.load("values");
      await context.sync();

      const headOk = first.values.some((row) => row.some((text) => text.includes(QUOC_HIEU)));
      const footOk = last.values.some((row) => row.some((text) => text.includes(ctx.document.footer.signerName.trim())));
      return headOk && footOk
        ? { status: "done" as const, kind: "ok" as const, text: `Đã dựng công văn (${tables.items.length} bảng); con trỏ ở chỗ nội dung.` }
        : { status: "done" as const, kind: "warning" as const, text: "Đã dựng nhưng đọc lại không thấy đủ quốc hiệu hoặc họ tên người ký; hãy kiểm tra tài liệu." };
    });
  },
};

/** Các bước tạo công văn: thiết lập trang, cài style, dựng nội dung, đánh số trang. */
export function letterSteps(): Array<PipelineStep<LetterContext>> {
  return [pageStep, styleStep, insertStep, pageNumberStep];
}

export function createLetter(ctx: LetterContext, onProgress?: (event: ProgressEvent) => void): Promise<PipelineReport> {
  return runPipeline(letterSteps(), ctx, undefined, onProgress);
}
