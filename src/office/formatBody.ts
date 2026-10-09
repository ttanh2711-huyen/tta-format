import { prepareParagraphOoxml } from "../ooxml/exactSpacing";
import { verifyExactPass } from "../ooxml/exactVerify";
import { buildStyleSpecs, type StyleSpec } from "../ooxml/styleSpecs";
import type { StandardProfile } from "../standards/types";
import { planFormatBody, summarizeFormat, type ParagraphInfo, type SkippedCounts } from "./paragraphClassifier";

export type FormatScope = "selection" | "document";
/** keep: giữ định dạng cục bộ (đậm, nghiêng, căn lề…); clear: xoá định dạng cục bộ để theo đúng style. */
export type LocalFormatting = "keep" | "clear";

export interface FormatBodyOptions {
  scope: FormatScope;
  localFormatting: LocalFormatting;
  /** Bộ style đang dùng (theo tuỳ chọn nhanh); bỏ trống thì theo mặc định của chuẩn. */
  specs?: readonly StyleSpec[];
}

export interface FormatBodyResult {
  applied: number;
  skipped: SkippedCounts;
  notes: string[];
  summary: { kind: "ok" | "warning"; text: string };
}

const CHUNK = 50;

/** Cùng mức với cài style (WordApi 1.6, Microsoft 365 bản 2308 trở lên). */
export function isFormatBodySupported(): boolean {
  return Office.context.requirements.isSetSupported("WordApi", "1.6");
}

function chunks<T>(items: readonly T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < items.length; i += size) {
    out.push(items.slice(i, i + size));
  }
  return out;
}

/**
 * Ghi lại từng đoạn bằng OOXML đã đặt giãn dòng Exactly. API Word không đặt được Exactly (spike T3),
 * nên đây là cách duy nhất. Mỗi đoạn được đọc OOXML rồi thay thế; bình luận hoặc đánh dấu gắn trong
 * đoạn có thể mất.
 */
export async function rewriteWithExact(context: Word.RequestContext, paragraphs: readonly Word.Paragraph[], pt: number): Promise<void> {
  for (const chunk of chunks(paragraphs, CHUNK)) {
    const ooxmls = chunk.map((paragraph) => paragraph.getOoxml());
    await context.sync();
    chunk.forEach((paragraph, index) => {
      paragraph.getRange("Whole").insertOoxml(prepareParagraphOoxml(ooxmls[index].value, pt), "Replace");
    });
    await context.sync();
  }
}

async function readTexts(context: Word.RequestContext, body: Word.Body): Promise<string[]> {
  const paragraphs = body.paragraphs;
  paragraphs.load("items/text");
  await context.sync();
  return paragraphs.items.map((paragraph) => paragraph.text);
}

/**
 * Phép thử cách ghi lại đoạn trên một tài liệu tạm (không mở ra, không đụng tài liệu của người dùng):
 * số đoạn, chữ và giãn dòng Exactly phải còn nguyên sau khi ghi. Trả về lý do nếu không đạt.
 */
async function selfTestExact(context: Word.RequestContext, pt: number): Promise<string | undefined> {
  const scratch = context.application.createDocument();
  scratch.body.insertText("Đoạn thử một\nĐoạn thử hai\nĐoạn thử ba", "Replace");
  await context.sync();
  const before = await readTexts(context, scratch.body);
  const paragraphs = scratch.body.paragraphs;
  paragraphs.load("items");
  await context.sync();
  await rewriteWithExact(context, paragraphs.items, pt);
  const after = await readTexts(context, scratch.body);
  const ooxml = scratch.body.getOoxml();
  await context.sync();
  const verdict = verifyExactPass({
    before,
    after,
    ooxml: ooxml.value,
    twips: Math.round(pt * 20),
    expectedExact: before.length,
  });
  return verdict.ok ? undefined : verdict.problem;
}

/**
 * Áp style nội dung của chuẩn lên các đoạn thường trong vùng chọn hoặc cả tài liệu. Bỏ qua đoạn trống,
 * đoạn trong bảng, đoạn đã dùng style của bộ chuẩn, tiêu đề có sẵn và mục danh sách tự động.
 *
 * Chuẩn có giãn dòng Exactly (Đảng) thì sau khi áp style, ghi lại từng đoạn bằng OOXML để đặt Exactly
 * trực tiếp trên đoạn; chỉ làm khi phép thử trên tài liệu tạm đạt, và kiểm tra lại sau khi ghi.
 */
export async function formatBody(profile: StandardProfile, options: FormatBodyOptions): Promise<FormatBodyResult> {
  if (!isFormatBodySupported()) {
    throw new Error("Định dạng nội dung cần Microsoft 365 bản 2308 trở lên (WordApi 1.6). Hãy cập nhật Word rồi mở lại.");
  }
  const specs = options.specs ?? buildStyleSpecs(profile);
  const styleName = `${profile.styles.prefix}NoiDung`;
  const spec = specs.find((item) => item.name === styleName);
  if (!spec) {
    throw new Error(`Chuẩn ${profile.name} không có style nội dung.`);
  }
  const exactPt = spec.paragraph.lineSpacing.kind === "exactly" ? spec.paragraph.lineSpacing.pt : undefined;

  return Word.run(async (context) => {
    const style = context.document.getStyles().getByNameOrNullObject(styleName);
    style.load("isNullObject");
    const scope = options.scope === "selection" ? context.document.getSelection().paragraphs : context.document.body.paragraphs;
    scope.load("items/text,items/style,items/styleBuiltIn,items/tableNestingLevel,items/isListItem");
    await context.sync();
    if (style.isNullObject) {
      throw new Error(`Chưa có style ${styleName}. Bấm "Cài bộ style" trước.`);
    }

    const infos: ParagraphInfo[] = scope.items.map((paragraph) => ({
      text: paragraph.text,
      styleName: paragraph.style,
      styleBuiltIn: paragraph.styleBuiltIn,
      tableNestingLevel: paragraph.tableNestingLevel,
      isListItem: paragraph.isListItem,
    }));
    const plan = planFormatBody(infos, profile);
    const targets = plan.targets.map((index) => scope.items[index]);

    const canResetParagraph = Office.context.requirements.isSetSupported("WordApiDesktop", "1.4");
    const canResetFont = Office.context.requirements.isSetSupported("WordApiDesktop", "1.3");
    const notes: string[] = [];
    let warn = false;

    for (const paragraph of targets) {
      paragraph.style = styleName;
      if (options.localFormatting === "clear") {
        if (canResetFont) paragraph.font.reset();
        if (canResetParagraph) paragraph.reset();
      }
    }
    await context.sync();
    if (options.localFormatting === "clear" && (!canResetFont || !canResetParagraph)) {
      notes.push("Chưa xoá hết định dạng cục bộ vì Word chưa hỗ trợ đủ (cần Microsoft 365 bản 2508 trở lên).");
      warn = true;
    }

    if (exactPt !== undefined && targets.length > 0) {
      const problem = await selfTestExact(context, exactPt);
      if (problem !== undefined) {
        notes.push(`Chưa đặt được giãn dòng Exactly ${exactPt} pt: phép thử trên tài liệu tạm không đạt (${problem}). Style đã áp, giãn dòng đang theo style (tối thiểu ${exactPt} pt).`);
        warn = true;
      } else {
        const before = await readTexts(context, context.document.body);
        await rewriteWithExact(context, targets, exactPt);
        const after = await readTexts(context, context.document.body);
        const ooxml = context.document.body.getOoxml();
        await context.sync();
        const verdict = verifyExactPass({
          before,
          after,
          ooxml: ooxml.value,
          twips: Math.round(exactPt * 20),
          expectedExact: targets.length,
        });
        if (verdict.ok) {
          notes.push(`Giãn dòng Exactly ${exactPt} pt đã đặt trên từng đoạn; bình luận hoặc đánh dấu gắn trong đoạn có thể không còn.`);
        } else {
          notes.push(`Đã ghi lại đoạn nhưng kiểm tra không đạt (${verdict.problem}). Hãy bấm Ctrl+Z để hoàn tác rồi kiểm tra tài liệu.`);
          warn = true;
        }
      }
    }

    const summary = summarizeFormat({ applied: targets.length, skipped: plan.skipped, notes, warn });
    return { applied: targets.length, skipped: plan.skipped, notes, summary };
  });
}

export interface FormatBodyPreview {
  targets: number;
  skipped: SkippedCounts;
  /** Chưa có style nội dung của chuẩn trong tài liệu. */
  styleMissing: boolean;
}

/** Chỉ đọc: sẽ có bao nhiêu đoạn được định dạng và bao nhiêu đoạn bị bỏ qua. */
export async function previewFormatBody(profile: StandardProfile, scope: FormatScope): Promise<FormatBodyPreview> {
  const styleName = `${profile.styles.prefix}NoiDung`;
  return Word.run(async (context) => {
    const style = context.document.getStyles().getByNameOrNullObject(styleName);
    style.load("isNullObject");
    const paragraphs = scope === "selection" ? context.document.getSelection().paragraphs : context.document.body.paragraphs;
    paragraphs.load("items/text,items/style,items/styleBuiltIn,items/tableNestingLevel,items/isListItem");
    await context.sync();
    const plan = planFormatBody(
      paragraphs.items.map((paragraph) => ({
        text: paragraph.text,
        styleName: paragraph.style,
        styleBuiltIn: paragraph.styleBuiltIn,
        tableNestingLevel: paragraph.tableNestingLevel,
        isListItem: paragraph.isListItem,
      })),
      profile,
    );
    return { targets: plan.targets.length, skipped: plan.skipped, styleMissing: style.isNullObject };
  });
}
