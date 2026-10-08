import { flatOpc } from "../ooxml/flatOpc";
import { isPlanEmpty, pageNumberParagraph, planPageNumbers, readPageNumberState, type PageNumberOptions, type SectionPlan, type SectionState } from "../ooxml/pageNumberHeader";

/** Đánh số trang cần WordApiDesktop 1.3 (đặt "trang đầu khác biệt" qua PageSetup). */
export function isPageNumberSupported(): boolean {
  return Office.context.requirements.isSetSupported("WordApiDesktop", "1.3");
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
  const states = sections.map((section, index) => {
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

    const after = planPageNumbers((await readStates(context)).states, options);
    const problems = after.flatMap((item, index) => {
      const where = sections.length > 1 ? `Phần ${index + 1}: ` : "";
      const found: string[] = [];
      if (item.writePrimary) found.push(`${where}header chưa có số trang đúng (căn giữa, ${options.sizePt} pt)`);
      if (item.setDifferentFirst !== undefined) found.push(`${where}chưa đặt được "trang đầu khác biệt"`);
      if (item.clearFirst) found.push(`${where}header trang đầu vẫn còn số trang`);
      return found;
    });
    return { sections: sections.length, changed: plan.filter((item) => !isPlanEmpty([item])).length, problems };
  });
}
