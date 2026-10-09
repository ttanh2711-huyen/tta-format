import { resolveParagraph, styleNameOf, type StyleTable } from "../ooxml/styleResolver";
import type { DocModel, ParagraphModel } from "../rules/docModel";
import type { CheckDocument, CheckParagraph, PageInfo } from "./types";

/**
 * Dựng tài liệu cho bộ kiểm tra từ mô hình OOXML: trải phẳng các khối theo thứ tự đọc (đoạn ngoài bảng, rồi
 * đoạn của từng ô theo hàng) để chỉ số khớp `body.paragraphs` của Word, và giải style cho từng đoạn.
 * Giới hạn: bảng lồng trong ô và khối `w:sdt` không được đọc nên chỉ số sau chúng có thể lệch.
 */
export function buildCheckDocument(model: DocModel, styles: StyleTable, pages: PageInfo[]): CheckDocument {
  const paragraphs: CheckParagraph[] = [];
  const add = (paragraph: ParagraphModel, inTable: boolean, table?: CheckParagraph["table"]): void => {
    const resolved = resolveParagraph(paragraph, styles);
    paragraphs.push({
      index: paragraphs.length,
      inTable,
      table,
      styleName: styleNameOf(paragraph.styleId ?? styles.defaultParagraphStyle ?? "", styles),
      resolved,
    });
  };
  let tableIndex = -1;
  for (const block of model.blocks) {
    if (block.kind === "paragraph") {
      add(block, false);
    } else {
      tableIndex += 1;
      block.rows.forEach((row, rowIndex) =>
        row.forEach((cell, colIndex) => {
          for (const paragraph of cell) add(paragraph, true, { index: tableIndex, row: rowIndex, col: colIndex });
        }),
      );
    }
  }
  return { pages, paragraphs };
}
