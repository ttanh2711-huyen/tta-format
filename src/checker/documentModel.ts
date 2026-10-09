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
  const add = (paragraph: ParagraphModel, inTable: boolean): void => {
    const resolved = resolveParagraph(paragraph, styles);
    paragraphs.push({
      index: paragraphs.length,
      inTable,
      styleName: styleNameOf(paragraph.styleId ?? styles.defaultParagraphStyle ?? "", styles),
      resolved,
    });
  };
  for (const block of model.blocks) {
    if (block.kind === "paragraph") {
      add(block, false);
    } else {
      for (const row of block.rows) for (const cell of row) for (const paragraph of cell) add(paragraph, true);
    }
  }
  return { pages, paragraphs };
}
