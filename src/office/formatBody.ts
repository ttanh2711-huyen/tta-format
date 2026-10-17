import { prepareParagraphOoxml } from "../ooxml/exactSpacing";
import { verifyExactPass } from "../ooxml/exactVerify";
import { buildStyleSpecs, type StyleSpec } from "../ooxml/styleSpecs";
import type { StandardProfile } from "../standards/types";
import { planFormatBody, summarizeFormat, type ParagraphInfo, type SkippedCounts } from "./paragraphClassifier";
import { CENTERED_ROLES } from "./paragraphRoles";

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

/** Tài liệu chưa có style nội dung của chuẩn: bảng Định dạng cài bộ style rồi chạy lại. */
export class StyleMissingError extends Error {}

export const NO_SELECTION = "Chưa tô đen nội dung. Hãy bôi đen các đoạn cần định dạng rồi bấm lại.";

/**
 * Phần tô đen: ghi thẳng phông, cỡ chữ, căn lề, thụt dòng, giãn đoạn, giãn dòng của style lên đoạn, để đoạn đúng chuẩn
 * dù trước đó có định dạng cục bộ khác (đậm, nghiêng, gạch chân trong đoạn giữ nguyên).
 */
function applyDirect(paragraph: Word.Paragraph, spec: StyleSpec): void {
  paragraph.font.name = spec.font.name;
  paragraph.font.size = spec.font.sizePt;
  paragraph.alignment = spec.paragraph.alignment;
  paragraph.leftIndent = 0;
  paragraph.firstLineIndent = spec.paragraph.firstLineIndentPt;
  paragraph.spaceBefore = spec.paragraph.spaceBeforePt;
  paragraph.spaceAfter = spec.paragraph.spaceAfterPt;
  if (spec.paragraph.lineSpacing.kind === "multiple") paragraph.lineSpacing = spec.paragraph.lineSpacing.lines * 12;
}

/** Đọc style, danh sách, bảng của đoạn cần WordApi 1.3 (Word 2016 trở lên). Kiểm tra style đã có dùng API style (WordApi 1.6). */
export function isFormatBodySupported(): boolean {
  return Office.context.requirements.isSetSupported("WordApi", "1.3");
}

const hasStyleApi = (): boolean => Office.context.requirements.isSetSupported("WordApi", "1.6");

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
    throw new Error("Định dạng nội dung cần Word 2016 trở lên. Hãy cập nhật Word rồi mở lại.");
  }
  const specs = options.specs ?? buildStyleSpecs(profile);
  const styleName = `${profile.styles.prefix}NoiDung`;
  const spec = specs.find((item) => item.name === styleName);
  if (!spec) {
    throw new Error(`Chuẩn ${profile.name} không có style nội dung.`);
  }
  const exactPt = spec.paragraph.lineSpacing.kind === "exactly" ? spec.paragraph.lineSpacing.pt : undefined;

  return Word.run(async (context) => {
    const style = hasStyleApi() ? context.document.getStyles().getByNameOrNullObject(styleName) : undefined;
    style?.load("isNullObject");
    const selected = options.scope === "selection";
    const selection = context.document.getSelection();
    if (selected) selection.load("text");
    const scope = selected ? selection.paragraphs : context.document.body.paragraphs;
    scope.load("items/text,items/style,items/styleBuiltIn,items/tableNestingLevel,items/isListItem");
    await context.sync();
    if (selected && selection.text.trim() === "") throw new Error(NO_SELECTION);
    if (style?.isNullObject) {
      throw new StyleMissingError(`Chưa có style ${styleName}. Bấm "Cài bộ style" trước.`);
    }

    const infos: ParagraphInfo[] = scope.items.map((paragraph) => ({
      text: paragraph.text,
      styleName: paragraph.style,
      styleBuiltIn: paragraph.styleBuiltIn,
      tableNestingLevel: paragraph.tableNestingLevel,
      isListItem: paragraph.isListItem,
    }));
    const plan = planFormatBody(infos, profile, { selection: selected });
    const targets = plan.targets.map((index) => scope.items[index]);
    // Style theo vai trò đoạn (issue #3); style nào chưa có trong tài liệu thì dùng style nội dung.
    const wanted = hasStyleApi() ? [...new Set(plan.roles.map((r) => r.styleName))].filter((name) => name !== styleName) : [];
    const roleStyles = wanted.map((name) => {
      const item = context.document.getStyles().getByNameOrNullObject(name);
      item.load("isNullObject");
      return { name, item };
    });
    await context.sync();
    const missing = new Set(roleStyles.filter(({ item }) => item.isNullObject).map(({ name }) => name));

    const canResetParagraph = Office.context.requirements.isSetSupported("WordApiDesktop", "1.4");
    const canResetFont = Office.context.requirements.isSetSupported("WordApiDesktop", "1.3");
    const notes: string[] = [];
    let warn = false;

    plan.roles.forEach((role, i) => {
      const paragraph = targets[i];
      const applied = missing.has(role.styleName) ? styleName : role.styleName;
      paragraph.style = applied;
      if (options.localFormatting === "clear") {
        if (canResetFont) paragraph.font.reset();
        if (canResetParagraph) paragraph.reset();
      }
      if (selected) applyDirect(paragraph, specs.find((item) => item.name === applied) ?? spec);
      if (role.bold) paragraph.font.bold = true;
    });
    await context.sync();
    const headings = plan.roles.filter((r) => r.role !== "noiDung" && r.role !== "y").length;
    if (headings > 0) notes.push(`Nhận ra ${headings} đề mục (chương, mục, điều, khoản, điểm, đề mục La Mã) và áp style đề mục tương ứng.`);
    if (missing.size > 0) {
      notes.push(`Chưa có style ${[...missing].join(", ")} nên các đoạn đó dùng style nội dung; bấm "Cài bộ style" rồi định dạng lại.`);
      warn = true;
    }
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
        // Tên chương, mục căn giữa không ghi lại giãn dòng Exactly như đoạn văn.
        const prose = targets.filter((_, i) => !CENTERED_ROLES.has(plan.roles[i].role));
        await rewriteWithExact(context, prose, exactPt);
        const after = await readTexts(context, context.document.body);
        const ooxml = context.document.body.getOoxml();
        await context.sync();
        const verdict = verifyExactPass({
          before,
          after,
          ooxml: ooxml.value,
          twips: Math.round(exactPt * 20),
          expectedExact: prose.length,
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
    const style = hasStyleApi() ? context.document.getStyles().getByNameOrNullObject(styleName) : undefined;
    style?.load("isNullObject");
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
    return { targets: plan.targets.length, skipped: plan.skipped, styleMissing: style?.isNullObject ?? false };
  });
}
