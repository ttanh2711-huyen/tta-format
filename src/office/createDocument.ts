import { BODY_MARKER, type BodyParagraph } from "../ooxml/documentAssembler";
import { flatOpc } from "../ooxml/flatOpc";
import { hasExactLineRule } from "../ooxml/facts";
import { styleIdFromName } from "../ooxml/styleXml";
import { pageNumberStep, pageStep, styleStep, type FixAllContext } from "../pipeline/steps";
import { runPipeline, type PipelineReport, type PipelineStep, type ProgressEvent } from "../pipeline/runner";

/** Ngữ cảnh của trình tạo văn bản: các tuỳ chọn của "Chỉnh toàn bộ" cộng với nội dung đã dựng sẵn của từng chuẩn. */
export interface LetterContext extends FixAllContext {
  /** Dựng thân văn bản (khối đầu, tên loại, chỗ nội dung, khối cuối) với đoạn chỗ nội dung đã cho. */
  buildBody: (body: BodyParagraph) => string;
  /** Chữ phải thấy trong bảng đầu tiên sau khi dựng (quốc hiệu hoặc tiêu đề Đảng). */
  headText: string;
  /** Họ tên người ký, phải thấy trong bảng cuối cùng. */
  signerName: string;
  /** Phần numbering.xml (cả thẻ `w:numbering`) nếu thân văn bản dùng đánh số tự động (Điều, Khoản, Điểm). */
  numberingXml?: string;
}

const insertStep: PipelineStep<LetterContext> = {
  id: "insert",
  label: "3. Dựng khối đầu, thân, khối cuối",
  order: 30,
  async plan() {
    return { willRun: true, lines: ["Thay toàn bộ nội dung bằng văn bản mới: khối đầu, các dòng thể loại hoặc Kính gửi, chỗ nội dung, khối cuối."] };
  },
  async run(ctx) {
    const bodyStyle = `${ctx.profile.styles.prefix}NoiDung`;
    const spec = ctx.specs.find((item) => item.name === bodyStyle);
    const exactPt = spec?.paragraph.lineSpacing.kind === "exactly" ? spec.paragraph.lineSpacing.pt : undefined;
    return Word.run(async (context) => {
      const body = context.document.body;
      body.insertOoxml(flatOpc({ bodyXml: ctx.buildBody({ styleId: styleIdFromName(bodyStyle), exactPt }), numberingXml: ctx.numberingXml }), "Replace");
      await context.sync();

      const paragraphs = body.paragraphs;
      paragraphs.load("items/text,items/style");
      const tables = body.tables;
      tables.load("items");
      await context.sync();
      const marker = paragraphs.items.find((paragraph) => paragraph.text.includes(BODY_MARKER));
      if (!marker) throw new Error("Không thấy chỗ đánh dấu nội dung sau khi chèn.");
      const notes: string[] = [];
      // Đoạn chỗ nội dung có thể kèm chữ dẫn đứng trước (ví dụ "Ban hành kèm theo Quyết định này "): chỉ xoá chữ đánh dấu.
      const hasLead = marker.text.trim() !== BODY_MARKER;
      if (hasLead) {
        const found = marker.search(BODY_MARKER);
        found.load("items");
        await context.sync();
        if (found.items.length === 0) throw new Error("Không tìm thấy chữ đánh dấu nội dung trong đoạn.");
        found.items[0].insertText("", "Replace");
      } else {
        marker.insertText("", "Replace");
      }
      if (marker.style !== bodyStyle) {
        marker.style = bodyStyle;
        if (exactPt !== undefined) notes.push(`Đoạn nội dung phải gán lại style ${bodyStyle} nên giãn dòng Exactly ${exactPt} pt có thể chưa còn; bấm "Định dạng nội dung" để đặt lại.`);
      }
      if (hasLead) marker.select("End");
      else marker.select();
      const markerXml = marker.getOoxml();
      const first = tables.items[0];
      const last = tables.items[tables.items.length - 1];
      first.load("values");
      last.load("values");
      await context.sync();

      if (exactPt !== undefined && !hasExactLineRule(markerXml.value, Math.round(exactPt * 20))) {
        notes.push(`Giãn dòng Exactly ${exactPt} pt chưa đặt được cho đoạn nội dung; bấm "Định dạng nội dung" sau khi gõ.`);
      }
      const headOk = first.values.some((row) => row.some((text) => text.includes(ctx.headText)));
      const footOk = last.values.some((row) => row.some((text) => text.includes(ctx.signerName.trim())));
      if (!headOk || !footOk) notes.push("Đọc lại không thấy đủ tiêu đề đầu văn bản hoặc họ tên người ký; hãy kiểm tra tài liệu.");
      return {
        status: "done" as const,
        kind: notes.length === 0 ? ("ok" as const) : ("warning" as const),
        text: [`Đã dựng văn bản (${tables.items.length} bảng); con trỏ ở chỗ nội dung.`, ...notes].join(" "),
      };
    });
  },
};

/** Các bước tạo văn bản: thiết lập trang, cài style, dựng nội dung, đánh số trang. */
export function letterSteps(): Array<PipelineStep<LetterContext>> {
  return [pageStep, styleStep, insertStep, pageNumberStep];
}

export function createLetter(ctx: LetterContext, onProgress?: (event: ProgressEvent) => void): Promise<PipelineReport> {
  return runPipeline(letterSteps(), ctx, undefined, onProgress);
}
