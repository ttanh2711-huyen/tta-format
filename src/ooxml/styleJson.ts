import type { LineSpacingSetting } from "../standards/types";
import type { StyleSpec } from "./styleSpecs";

/** Word.ParagraphFormat.lineSpacing tính bằng pt, trong đó 12 pt là một dòng đơn. */
export function lineSpacingPoints(setting: LineSpacingSetting): number {
  return setting.kind === "multiple" ? Math.round(setting.lines * 12 * 100) / 100 : setting.pt;
}

function toStyle(spec: StyleSpec): Record<string, unknown> {
  return {
    baseStyle: "Normal",
    builtIn: false,
    inUse: true,
    linked: false,
    nameLocal: spec.name,
    priority: 1,
    quickStyle: true,
    type: "Paragraph",
    unhideWhenUsed: false,
    visibility: false,
    paragraphFormat: {
      alignment: spec.paragraph.alignment,
      firstLineIndent: spec.paragraph.firstLineIndentPt,
      keepTogether: false,
      keepWithNext: spec.paragraph.keepWithNext,
      leftIndent: 0,
      lineSpacing: lineSpacingPoints(spec.paragraph.lineSpacing),
      lineUnitAfter: 0,
      lineUnitBefore: 0,
      mirrorIndents: false,
      outlineLevel: "OutlineLevelBodyText",
      rightIndent: 0,
      spaceAfter: spec.paragraph.spaceAfterPt,
      spaceBefore: spec.paragraph.spaceBeforePt,
      widowControl: true,
    },
    font: {
      name: spec.font.name,
      size: spec.font.sizePt,
      bold: spec.font.bold,
      italic: spec.font.italic,
      color: spec.font.color,
      underline: spec.font.underline ? "Single" : "None",
      subscript: false,
      superscript: false,
      strikeThrough: false,
      doubleStrikeThrough: false,
      highlightColor: null,
      hidden: false,
    },
  };
}

/**
 * Chuỗi JSON cho `document.importStylesFromJson`. Cách này đã được kiểm chứng ở spike T3 (cài mới và
 * cập nhật bằng Overwrite), nhưng **không đặt được giãn dòng Exactly** (Word cho ra quy tắc atLeast).
 */
export function toStyleImportJson(specs: readonly StyleSpec[]): string {
  return `${JSON.stringify({ styles: specs.map(toStyle) }, null, 2)}\n`;
}
