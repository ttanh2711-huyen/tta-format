import { flatOpc } from "../ooxml/flatOpc";
import { describePageNumberState, isPlanEmpty, pageNumberParagraph, pageNumberParagraphHideFirst, planPageNumbers, readPageNumberState, type PageNumberOptions, type SectionPlan, type SectionState } from "../ooxml/pageNumberHeader";

/**
 * Đánh số trang: có WordApiDesktop 1.3 (Microsoft 365) thì đặt "trang đầu khác biệt" qua PageSetup; Word 2016, 2019,
 * 2021 thì ghi header bằng trường IF ẩn số ở trang 1 (chỉ cần WordApi 1.3).
 */
export function isPageNumberSupported(): boolean {
  return Office.context.requirements.isSetSupported("WordApi", "1.3");
}

const hasPageSetupApi = (): boolean => Office.context.requirements.isSetSupported("WordApiDesktop", "1.3");

/** Header đã có số trang đúng cỡ, căn giữa thì không ghi lại. */
const primaryOk = (state: SectionState["primary"], options: PageNumberOptions): boolean => state.hasPage && state.centered && state.sizePt === options.sizePt;

async function applyLegacy(context: Word.RequestContext, options: PageNumberOptions): Promise<PageNumberResult> {
  const collection = context.document.sections;
  collection.load("items");
  await context.sync();
  const headers = collection.items.map((section) => section.getHeader("Primary"));
  const reads = headers.map((header) => header.getOoxml());
  await context.sync();
  let changed = 0;
  headers.forEach((header, index) => {
    if (primaryOk(readPageNumberState(reads[index].value), options)) return;
    // Chỉ phần đầu ẩn số ở trang 1, như bản Microsoft 365.
    const paragraph = index === 0 && options.hideFirst ? pageNumberParagraphHideFirst(options.sizePt) : pageNumberParagraph(options.sizePt);
    header.insertOoxml(flatOpc({ bodyXml: paragraph }), "Replace");
    changed += 1;
  });
  await context.sync();
  const after = headers.map((header) => header.getOoxml());
  await context.sync();
  const problems = after.flatMap((read, index) => (primaryOk(readPageNumberState(read.value), options) ? [] : [`${headers.length > 1 ? `Phần ${index + 1}: ` : ""}header chưa có số trang đúng (đọc lại: ${describePageNumberState(readPageNumberState(read.value))})`]));
  return { sections: headers.length, changed, problems };
}

async function readStates(context: Word.RequestContext): Promise<{ sections: Word.Section[]; states: SectionState[] }> {
  const collection = context.document.sections;
  collection.load("items");
  await context.sync();
  const sections = collection.items;
  const reads = sections.map((section) => {
    section.pageSetup.load("differentFirstPageHeaderFooter");
    const first = section.getHeader("FirstPage");
    first.load("text");
    return { primary: section.getHeader("Primary").getOoxml(), first, firstXml: first.getOoxml() };
  });
  await context.sync();
  const states: SectionState[] = sections.map((section, index) => {
    const read = reads[index];
    return {
      differentFirst: section.pageSetup.differentFirstPageHeaderFooter,
      primary: readPageNumberState(read.primary.value),
      first: { hasPage: readPageNumberState(read.firstXml.value).hasPage, empty: read.first.text.trim() === "" },
    };
  });
  return { sections, states };
}

/** Chỉ đọc: việc cần làm cho từng phần của tài liệu. */
export async function planPageNumberDocument(options: PageNumberOptions): Promise<SectionPlan[]> {
  if (!hasPageSetupApi()) {
    return Word.run(async (context) => {
      const collection = context.document.sections;
      collection.load("items");
      await context.sync();
      const reads = collection.items.map((section) => section.getHeader("Primary").getOoxml());
      await context.sync();
      return reads.map((read) => ({ setDifferentFirst: undefined, writePrimary: !primaryOk(readPageNumberState(read.value), options), clearFirst: false }));
    });
  }
  return Word.run(async (context) => planPageNumbers((await readStates(context)).states, options));
}

export interface PageNumberResult {
  sections: number;
  /** Số phần đã được sửa. */
  changed: number;
  /** Điểm chưa đúng sau khi ghi và đọc lại; rỗng nếu đã đúng. */
  problems: string[];
}

/** Áp dụng đánh số trang cho mọi phần rồi đọc lại để kiểm tra. Idempotent. */
export async function applyPageNumbers(options: PageNumberOptions): Promise<PageNumberResult> {
  if (!hasPageSetupApi()) return Word.run((context) => applyLegacy(context, options));
  return Word.run(async (context) => {
    const { sections, states } = await readStates(context);
    const plan = planPageNumbers(states, options);
    plan.forEach((item, index) => {
      const section = sections[index];
      if (item.setDifferentFirst !== undefined) section.pageSetup.differentFirstPageHeaderFooter = item.setDifferentFirst;
      if (item.clearFirst) section.getHeader("FirstPage").clear();
      if (item.writePrimary) section.getHeader("Primary").insertOoxml(flatOpc({ bodyXml: pageNumberParagraph(options.sizePt) }), "Replace");
    });
    await context.sync();

    // insertOoxml có thể làm mất thuộc tính của đoạn (căn giữa, giãn đoạn; quy tắc đoạn đơn của spike T3) và
    // giữ lại đoạn trống cũ: đặt lại bằng API và bỏ đoạn trống thừa.
    const written = plan.flatMap((item, index) => (item.writePrimary ? [sections[index].getHeader("Primary")] : []));
    const lists = written.map((header) => {
      header.paragraphs.load("items/text");
      return header.paragraphs;
    });
    await context.sync();
    lists.forEach((list) => {
      const withText = list.items.filter((paragraph) => paragraph.text.trim() !== "");
      const keep = withText.length > 0 ? withText : list.items.slice(0, 1);
      for (const paragraph of list.items) {
        if (!keep.includes(paragraph) && list.items.length > 1) paragraph.delete();
      }
      for (const paragraph of keep) {
        paragraph.alignment = "Centered";
        paragraph.spaceBefore = 0;
        paragraph.spaceAfter = 0;
      }
    });
    await context.sync();

    const afterStates = (await readStates(context)).states;
    const after = planPageNumbers(afterStates, options);
    const problems = after.flatMap((item, index) => {
      const where = sections.length > 1 ? `Phần ${index + 1}: ` : "";
      const found: string[] = [];
      if (item.writePrimary) found.push(`${where}header chưa có số trang đúng (cần căn giữa, ${options.sizePt} pt; đọc lại: ${describePageNumberState(afterStates[index].primary)})`);
      if (item.setDifferentFirst !== undefined) found.push(`${where}chưa đặt được "trang đầu khác biệt"`);
      if (item.clearFirst) found.push(`${where}header trang đầu vẫn còn số trang`);
      return found;
    });
    return { sections: sections.length, changed: plan.filter((item) => !isPlanEmpty([item])).length, problems };
  });
}
