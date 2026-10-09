/**
 * Mô hình định dạng của tài liệu, đọc từ OOXML: đủ để so với fixture (T2A) và để bộ kiểm tra (T13) dùng lại.
 * Chỉ đọc định dạng trực tiếp (không giải style) và chỉ các phần tử cấp thân (bảng và đoạn); khối `w:sdt` bị bỏ qua.
 */

export interface RunProps {
  fontName?: string;
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
}

export interface ParagraphModel extends RunProps {
  kind: "paragraph";
  text: string;
  align?: "left" | "center" | "right" | "both";
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  lineMultiple?: number;
  lineExactPt?: number;
  lineAtLeastPt?: number;
  firstLineIndentMm?: number;
  leftIndentMm?: number;
  /** Thụt treo (mm) của dòng gạch đầu dòng. */
  hangingIndentMm?: number;
  /** Mã style của đoạn (w:pStyle); đoạn không khai báo thì dùng style mặc định của tài liệu. */
  styleId?: string;
  /** Độ dài (mm) của các đường thẳng nổi neo trong đoạn. */
  floatingLinesMm?: number[];
  borderBottom?: boolean;
}

export interface TableModel {
  kind: "table";
  widthMm?: number;
  indentMm?: number;
  columnsMm: number[];
  /** Mọi đường viền của bảng được khai báo rõ là không có. */
  bordersNone: boolean;
  /** hàng → ô → các đoạn trong ô. */
  rows: ParagraphModel[][][];
}

export interface PageModel {
  widthMm: number;
  heightMm: number;
  marginsMm: { top: number; bottom: number; left: number; right: number };
  headerDistanceMm: number;
  footerDistanceMm: number;
  /** Trang đầu khác biệt. */
  titlePg: boolean;
  /** Có trường số trang PAGE trong một header. */
  pageNumberInHeader: boolean;
}

export type Block = ParagraphModel | TableModel;

export interface DocModel {
  page: PageModel;
  blocks: Block[];
}

export interface DocParts {
  documentXml: string;
  headerXmls?: string[];
}

const TWIP_MM = 25.4 / 1440;
const EMU_MM = 1 / 36000;

const round = (value: number, digits: number): number => Math.round(value * 10 ** digits) / 10 ** digits;
const twipsToMm = (twips: number): number => round(twips * TWIP_MM, 1);

function attr(tag: string, name: string): string | undefined {
  return new RegExp(`\\bw:${name}="([^"]*)"`).exec(tag)?.[1];
}

function firstTag(xml: string, name: string): string | undefined {
  return new RegExp(`<w:${name}(?=[\\s>/])[^>]*>`).exec(xml)?.[0];
}

function decode(text: string): string {
  return text.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
}

/** Vị trí kết thúc của phần tử `w:tag` bắt đầu tại `start`, có tính phần tử cùng tên lồng bên trong. */
function elementEnd(xml: string, start: number, tag: string): number {
  const gt = xml.indexOf(">", start);
  if (gt < 0) return xml.length;
  if (xml[gt - 1] === "/") return gt + 1;
  const token = new RegExp(`<(/?)w:${tag}(?=[\\s>/])[^>]*?(/?)>`, "g");
  token.lastIndex = gt + 1;
  let depth = 1;
  let match: RegExpExecArray | null;
  while (depth > 0 && (match = token.exec(xml)) !== null) {
    if (match[1] === "/") depth -= 1;
    else if (match[2] !== "/") depth += 1;
  }
  return depth === 0 ? token.lastIndex : xml.length;
}

/** Các phần tử `w:tag` ở cấp ngoài cùng của `xml`. */
function elements(xml: string, tag: string): string[] {
  const open = new RegExp(`<w:${tag}(?=[\\s>/])`, "g");
  const out: string[] = [];
  let match: RegExpExecArray | null;
  while ((match = open.exec(xml)) !== null) {
    const end = elementEnd(xml, match.index, tag);
    out.push(xml.slice(match.index, end));
    open.lastIndex = end;
  }
  return out;
}

function readRunProps(rPr: string): RunProps {
  const props: RunProps = {};
  const fonts = firstTag(rPr, "rFonts");
  const name = fonts ? (attr(fonts, "ascii") ?? attr(fonts, "hAnsi")) : undefined;
  if (name) props.fontName = name;
  const size = firstTag(rPr, "sz");
  if (size) props.sizePt = Number(attr(size, "val")) / 2;
  const flag = (tagName: string): boolean | undefined => {
    const tag = firstTag(rPr, tagName);
    if (!tag) return undefined;
    const value = attr(tag, "val");
    return value !== "0" && value !== "false" && value !== "none";
  };
  const bold = flag("b");
  if (bold !== undefined) props.bold = bold;
  const italic = flag("i");
  if (italic !== undefined) props.italic = italic;
  const underline = flag("u");
  if (underline !== undefined) props.underline = underline;
  return props;
}

function stripFallback(xml: string): string {
  return xml.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, "");
}

export function readParagraph(raw: string): ParagraphModel {
  const xml = stripFallback(raw);
  const model: ParagraphModel = { kind: "paragraph", text: "" };
  const pPr = /<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(xml)?.[1] ?? "";
  // Phần dấu đoạn (rPr trong pPr) tách riêng để không lẫn với thuộc tính của run.
  const markRPr = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(pPr)?.[1] ?? "";
  const pPrOnly = pPr.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, "");

  const styleTag = firstTag(pPrOnly, "pStyle");
  const styleId = styleTag ? attr(styleTag, "val") : undefined;
  if (styleId) model.styleId = styleId;

  const jc = firstTag(pPrOnly, "jc");
  const align = jc ? attr(jc, "val") : undefined;
  if (align === "left" || align === "center" || align === "right" || align === "both") model.align = align;

  const spacing = firstTag(pPrOnly, "spacing");
  if (spacing) {
    const before = attr(spacing, "before");
    const after = attr(spacing, "after");
    const line = attr(spacing, "line");
    const rule = attr(spacing, "lineRule") ?? "auto";
    if (before !== undefined) model.spaceBeforePt = Number(before) / 20;
    if (after !== undefined) model.spaceAfterPt = Number(after) / 20;
    if (line !== undefined) {
      if (rule === "exact") model.lineExactPt = Number(line) / 20;
      else if (rule === "atLeast") model.lineAtLeastPt = Number(line) / 20;
      else model.lineMultiple = round(Number(line) / 240, 2);
    }
  }

  const ind = firstTag(pPrOnly, "ind");
  if (ind) {
    const first = attr(ind, "firstLine");
    const left = attr(ind, "left") ?? attr(ind, "start");
    if (first !== undefined) model.firstLineIndentMm = twipsToMm(Number(first));
    if (left !== undefined) model.leftIndentMm = twipsToMm(Number(left));
    const hanging = attr(ind, "hanging");
    if (hanging !== undefined) model.hangingIndentMm = twipsToMm(Number(hanging));
  }

  if (/<w:pBdr>[\s\S]*?<w:bottom\b(?![^>]*w:val="(?:none|nil)")[^>]*>/.test(pPrOnly)) model.borderBottom = true;

  const lines = [...xml.matchAll(/<wp:anchor\b[\s\S]*?<\/wp:anchor>/g)]
    .filter((anchor) => /prst="line"/.test(anchor[0]))
    .map((anchor) => round(Number(/<wp:extent\b[^>]*\bcx="(\d+)"/.exec(anchor[0])?.[1] ?? 0) * EMU_MM, 1));
  if (lines.length > 0) model.floatingLinesMm = lines;

  const runs = elements(xml.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, ""), "r");
  let text = "";
  let runProps: string | undefined;
  for (const run of runs) {
    const piece = [...run.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\s*\/>/g)].map((m) => (m[1] === undefined ? "\t" : decode(m[1]))).join("");
    if (piece !== "" && runProps === undefined) runProps = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(run)?.[1] ?? "";
    text += piece;
  }
  model.text = text;
  Object.assign(model, readRunProps(runProps ?? markRPr));
  return model;
}

function readTable(xml: string): TableModel {
  const tblPr = /<w:tblPr>([\s\S]*?)<\/w:tblPr>/.exec(xml)?.[1] ?? "";
  const width = firstTag(tblPr, "tblW");
  const indent = firstTag(tblPr, "tblInd");
  const borders = /<w:tblBorders>([\s\S]*?)<\/w:tblBorders>/.exec(tblPr)?.[1];
  const sides = borders ? [...borders.matchAll(/<w:(?:top|left|bottom|right|insideH|insideV)\b[^>]*>/g)].map((m) => attr(m[0], "val")) : [];
  const grid = [...(/<w:tblGrid>([\s\S]*?)<\/w:tblGrid>/.exec(xml)?.[1] ?? "").matchAll(/<w:gridCol\b[^>]*>/g)].map((m) => twipsToMm(Number(attr(m[0], "w"))));
  const table: TableModel = {
    kind: "table",
    columnsMm: grid,
    bordersNone: sides.length > 0 && sides.every((value) => value === "none" || value === "nil"),
    rows: elements(xml.replace(/<w:tblPr>[\s\S]*?<\/w:tblPr>/, ""), "tr").map((row) =>
      elements(row, "tc").map((cell) => elements(cell.replace(/<w:tcPr>[\s\S]*?<\/w:tcPr>/, ""), "p").map(readParagraph)),
    ),
  };
  if (width && attr(width, "type") === "dxa") table.widthMm = twipsToMm(Number(attr(width, "w")));
  if (indent) table.indentMm = twipsToMm(Number(attr(indent, "w")));
  return table;
}

function readBlocks(body: string): Block[] {
  const blocks: Block[] = [];
  const open = /<w:(tbl|p)(?=[\s>/])/g;
  let match: RegExpExecArray | null;
  while ((match = open.exec(body)) !== null) {
    const end = elementEnd(body, match.index, match[1]);
    const xml = body.slice(match.index, end);
    blocks.push(match[1] === "tbl" ? readTable(xml) : readParagraph(xml));
    open.lastIndex = end;
  }
  return blocks;
}

function readPage(documentXml: string, headerXmls: readonly string[]): PageModel {
  const sections = [...documentXml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)];
  const sect = sections.length > 0 ? sections[sections.length - 1][0] : "";
  const size = firstTag(sect, "pgSize") ?? firstTag(sect, "pgSz") ?? "";
  const mar = firstTag(sect, "pgMar") ?? "";
  const mm = (tag: string, name: string): number => twipsToMm(Number(attr(tag, name) ?? 0));
  return {
    widthMm: mm(size, "w"),
    heightMm: mm(size, "h"),
    marginsMm: { top: mm(mar, "top"), bottom: mm(mar, "bottom"), left: mm(mar, "left"), right: mm(mar, "right") },
    headerDistanceMm: mm(mar, "header"),
    footerDistanceMm: mm(mar, "footer"),
    titlePg: /<w:titlePg\b(?![^>]*w:val="(?:0|false)")/.test(sect),
    pageNumberInHeader: headerXmls.some((xml) => /<w:instrText\b[^>]*>[^<]*\bPAGE\b/.test(xml) || /<w:fldSimple\b[^>]*w:instr="[^"]*\bPAGE\b/.test(xml)),
  };
}

export function extractDocModel(parts: DocParts): DocModel {
  const body = /<w:body>([\s\S]*)<\/w:body>/.exec(parts.documentXml)?.[1] ?? "";
  return { page: readPage(parts.documentXml, parts.headerXmls ?? []), blocks: readBlocks(body.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>\s*$/, "")) };
}

/** Chuỗi công khai của thể thức, giữ nguyên khi ẩn danh. */
const PUBLIC_TEXTS: ReadonlySet<string> = new Set(["CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc", "ĐẢNG CỘNG SẢN VIỆT NAM", "Nơi nhận:", "STT"]);
/** Nhãn đầu dòng được giữ, phần sau bị thay. */
const LABELS: readonly string[] = ["Kính gửi:", "Kính trình:", "Số:", "V/v", "Lưu:", "Căn cứ"];

function anonymizeText(text: string): string {
  const trimmed = text.trim();
  if (trimmed === "") return text;
  if (PUBLIC_TEXTS.has(trimmed)) return trimmed;
  const label = LABELS.find((candidate) => trimmed.startsWith(candidate));
  return label ? `${label} […]` : "[…]";
}

function anonymizeParagraph(p: ParagraphModel): ParagraphModel {
  return { ...p, text: anonymizeText(p.text) };
}

/** Bản sao đã ẩn danh: giữ nguyên toàn bộ định dạng, chỉ chữ thể thức công khai và nhãn đầu dòng còn lại. */
export function anonymizeModel(model: DocModel): DocModel {
  return {
    page: model.page,
    blocks: model.blocks.map((block) =>
      block.kind === "paragraph" ? anonymizeParagraph(block) : { ...block, rows: block.rows.map((row) => row.map((cell) => cell.map(anonymizeParagraph))) },
    ),
  };
}
