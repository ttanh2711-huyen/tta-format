import type { StandardProfile } from "../standards/types";
import { classifyTable, type TableKind } from "./paragraphClassifier";

export interface TableSurvey {
  data: number;
  layout: number;
  nested: number;
  /** Số đoạn trong các bảng dữ liệu chưa dùng style chữ bảng của chuẩn. */
  paragraphsToStyle: number;
  styleMissing: boolean;
}

const BORDERS = ["Top", "Bottom", "Left", "Right", "InsideHorizontal", "InsideVertical"] as const;

interface SurveyedTable {
  kind: TableKind;
  table: Word.Table;
  paragraphs: Word.Paragraph[];
}

async function survey(context: Word.RequestContext, profile: StandardProfile): Promise<{ tables: SurveyedTable[]; result: TableSurvey }> {
  const styleName = `${profile.styles.prefix}BangDuLieu`;
  // Kiểm tra style đã có dùng API style (WordApi 1.6); Word 2016, 2019, 2021 thì coi như đã cài ở bước trước.
  const style = Office.context.requirements.isSetSupported("WordApi", "1.6") ? context.document.getStyles().getByNameOrNullObject(styleName) : undefined;
  style?.load("isNullObject");
  const tables = context.document.body.tables;
  tables.load("items/nestingLevel");
  await context.sync();
  const borders = tables.items.map((table) => BORDERS.map((location) => table.getBorder(location)));
  for (const row of borders) for (const border of row) border.load("type");
  await context.sync();

  const kinds = tables.items.map((table, index) =>
    classifyTable({ nestingLevel: table.nestingLevel, borderTypes: borders[index].map((border) => String(border.type)) }),
  );
  const dataTables = tables.items.filter((_, index) => kinds[index] === "data");
  const paragraphSets = dataTables.map((table) => {
    const paragraphs = table.getRange("Whole").paragraphs;
    paragraphs.load("items/style");
    return paragraphs;
  });
  await context.sync();

  const surveyed: SurveyedTable[] = dataTables.map((table, index) => ({
    kind: "data" as const,
    table,
    paragraphs: paragraphSets[index].items.filter((paragraph) => paragraph.style !== styleName),
  }));
  return {
    tables: surveyed,
    result: {
      data: dataTables.length,
      layout: kinds.filter((kind) => kind === "layout").length,
      nested: kinds.filter((kind) => kind === "nested").length,
      paragraphsToStyle: surveyed.reduce((sum, item) => sum + item.paragraphs.length, 0),
      styleMissing: style?.isNullObject ?? false,
    },
  };
}

/** Chỉ đọc: có bao nhiêu bảng dữ liệu, bảng bố cục và đoạn cần áp style chữ bảng. */
export async function surveyTables(profile: StandardProfile): Promise<TableSurvey> {
  return Word.run(async (context) => (await survey(context, profile)).result);
}

/**
 * Áp style chữ bảng của chuẩn cho các đoạn trong bảng dữ liệu (bảng có viền). Bảng không viền, nơi
 * thường đặt đầu văn bản, nơi nhận và chữ ký, không bị đụng tới. Viền và hàng tiêu đề không đổi.
 */
export async function formatTables(profile: StandardProfile, localFormatting: "keep" | "clear"): Promise<TableSurvey> {
  const styleName = `${profile.styles.prefix}BangDuLieu`;
  return Word.run(async (context) => {
    const { tables, result } = await survey(context, profile);
    if (result.styleMissing) {
      throw new Error(`Chưa có style ${styleName}. Bước cài style phải chạy trước.`);
    }
    const canResetFont = Office.context.requirements.isSetSupported("WordApiDesktop", "1.3");
    const canResetParagraph = Office.context.requirements.isSetSupported("WordApiDesktop", "1.4");
    for (const { paragraphs } of tables) {
      for (const paragraph of paragraphs) {
        paragraph.style = styleName;
        if (localFormatting === "clear") {
          if (canResetFont) paragraph.font.reset();
          if (canResetParagraph) paragraph.reset();
        }
      }
    }
    await context.sync();
    return result;
  });
}
