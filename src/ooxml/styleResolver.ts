import { readParagraph, type ParagraphModel } from "../rules/docModel";

/**
 * Giải style để có định dạng hiệu lực của đoạn: mặc định của tài liệu, rồi chuỗi style (basedOn), rồi định dạng
 * trực tiếp trên đoạn. Chỉ xét style đoạn; style ký tự (rStyle) và danh sách đánh số chưa được xét.
 */

interface StyleEntry {
  name: string;
  basedOn?: string;
  props: ParagraphModel;
}

export interface StyleTable {
  defaults: ParagraphModel;
  defaultParagraphStyle?: string;
  styles: Map<string, StyleEntry>;
}

function attrOf(tag: string | undefined, name: string): string | undefined {
  return tag ? new RegExp(`\\bw:${name}="([^"]*)"`).exec(tag)?.[1] : undefined;
}

function propsFrom(pPrInner: string, rPrInner: string): ParagraphModel {
  const pPr = pPrInner.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, "");
  return readParagraph(`<w:p><w:pPr>${pPr}<w:rPr>${rPrInner}</w:rPr></w:pPr></w:p>`);
}

function splitStyle(block: string): { pPr: string; rPr: string } {
  const pPr = /<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(block)?.[1] ?? "";
  const rest = block.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, "");
  return { pPr, rPr: /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(rest)?.[1] ?? "" };
}

export function parseStyleTable(stylesXml: string): StyleTable {
  const defaultsBlock = /<w:docDefaults>([\s\S]*?)<\/w:docDefaults>/.exec(stylesXml)?.[1] ?? "";
  const defaults = propsFrom(
    /<w:pPrDefault>\s*<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(defaultsBlock)?.[1] ?? "",
    /<w:rPrDefault>\s*<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(defaultsBlock)?.[1] ?? "",
  );
  const styles = new Map<string, StyleEntry>();
  let defaultParagraphStyle: string | undefined;
  for (const match of stylesXml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
    const [, attrs, body] = match;
    if (attrOf(attrs, "type") !== "paragraph") continue;
    const id = attrOf(attrs, "styleId");
    if (!id) continue;
    if (attrOf(attrs, "default") === "1") defaultParagraphStyle = id;
    const { pPr, rPr } = splitStyle(body);
    styles.set(id, {
      name: attrOf(/<w:name\b[^>]*>/.exec(body)?.[0], "val") ?? id,
      basedOn: attrOf(/<w:basedOn\b[^>]*>/.exec(body)?.[0], "val"),
      props: propsFrom(pPr, rPr),
    });
  }
  return { defaults, defaultParagraphStyle, styles };
}

export function styleNameOf(styleId: string, table: StyleTable): string | undefined {
  return table.styles.get(styleId)?.name;
}

const LINE_KEYS = ["lineMultiple", "lineExactPt", "lineAtLeastPt"] as const;
/** Các khoá thuộc về chính đoạn, không phải định dạng kế thừa. */
const OWN_KEYS = new Set(["kind", "text", "styleId", "floatingLinesMm", "borderBottom"]);

/** Gộp `over` lên `base`: khoá nào `over` có thì thắng; quy tắc giãn dòng là một nhóm loại trừ nhau. */
function merge(base: ParagraphModel, over: ParagraphModel): ParagraphModel {
  const out: Record<string, unknown> = { ...base };
  if (LINE_KEYS.some((key) => over[key] !== undefined)) for (const key of LINE_KEYS) delete out[key];
  for (const [key, value] of Object.entries(over)) {
    if (value !== undefined && !OWN_KEYS.has(key)) out[key] = value;
  }
  return out as unknown as ParagraphModel;
}

function chainOf(styleId: string | undefined, table: StyleTable): StyleEntry[] {
  const chain: StyleEntry[] = [];
  const seen = new Set<string>();
  let current = styleId ?? table.defaultParagraphStyle;
  while (current !== undefined && !seen.has(current)) {
    seen.add(current);
    const entry = table.styles.get(current);
    if (!entry) break;
    chain.unshift(entry);
    current = entry.basedOn;
  }
  return chain;
}

/** Đoạn với định dạng hiệu lực đã giải style; giữ nguyên chữ, mã style và các đặc điểm riêng của đoạn. */
export function resolveParagraph(paragraph: ParagraphModel, table: StyleTable): ParagraphModel {
  let effective: ParagraphModel = table.defaults;
  for (const entry of chainOf(paragraph.styleId, table)) effective = merge(effective, entry.props);
  effective = merge(effective, paragraph);
  return { ...effective, kind: "paragraph", text: paragraph.text, styleId: paragraph.styleId, floatingLinesMm: paragraph.floatingLinesMm, borderBottom: paragraph.borderBottom } as ParagraphModel;
}
