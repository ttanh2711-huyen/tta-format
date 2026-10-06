/**
 * Đọc dữ kiện từ chuỗi OOXML do Word trả về (getOoxml) để kiểm chứng thay đổi đã thật sự
 * vào tài liệu. Dùng biểu thức chính quy vì chỉ cần vài thuộc tính, không dựng lại cây XML.
 */

function attr(tag: string, name: string): string | undefined {
  const match = new RegExp(`\\bw:${name}="([^"]*)"`).exec(tag);
  return match ? match[1] : undefined;
}

function firstTag(xml: string, name: string): string | undefined {
  const match = new RegExp(`<w:${name}\\b[^>]*>`).exec(xml);
  return match ? match[0] : undefined;
}

function isNoBorder(value: string | undefined): boolean {
  return value === "none" || value === "nil";
}

/** Đoạn có viền dưới (khác "none", "nil"). */
export function hasBottomBorder(xml: string): boolean {
  const blocks = xml.match(/<w:pBdr>[\s\S]*?<\/w:pBdr>/g) ?? [];
  return blocks.some((block) => {
    const bottom = firstTag(block, "bottom");
    return bottom !== undefined && !isNoBorder(attr(bottom, "val"));
  });
}

export function readParagraphIndentTwips(xml: string): { left?: number; right?: number } {
  const tag = firstTag(xml, "ind");
  if (!tag) {
    return {};
  }
  const left = attr(tag, "left") ?? attr(tag, "start");
  const right = attr(tag, "right") ?? attr(tag, "end");
  const result: { left?: number; right?: number } = {};
  if (left !== undefined) result.left = Number(left);
  if (right !== undefined) result.right = Number(right);
  return result;
}

/** Cả 6 cạnh của bảng khai báo rõ là không viền. */
export function tableBordersAllNone(xml: string): boolean {
  const block = /<w:tblBorders>([\s\S]*?)<\/w:tblBorders>/.exec(xml);
  if (!block) {
    return false;
  }
  const sides: Array<string[]> = [["top"], ["left", "start"], ["bottom"], ["right", "end"], ["insideH"], ["insideV"]];
  return sides.every((names) =>
    names.some((name) => {
      const tag = firstTag(block[1], name);
      return tag !== undefined && isNoBorder(attr(tag, "val"));
    }),
  );
}

export function readTableIndentTwips(xml: string): number | undefined {
  const tag = firstTag(xml, "tblInd");
  const value = tag ? attr(tag, "w") : undefined;
  return value === undefined ? undefined : Number(value);
}

/** Có giãn dòng kiểu Exactly; nếu truyền `twips` thì phải đúng giá trị đó. */
export function hasExactLineRule(xml: string, twips?: number): boolean {
  const tags = xml.match(/<w:spacing\b[^>]*>/g) ?? [];
  return tags.some((tag) => {
    if (attr(tag, "lineRule") !== "exact") {
      return false;
    }
    return twips === undefined || Number(attr(tag, "line")) === twips;
  });
}

/** Có trường số trang PAGE (không tính NUMPAGES, PAGEREF). */
export function hasPageField(xml: string): boolean {
  return /<w:fldSimple\b[^>]*w:instr="[^"]*\bPAGE\b/.test(xml) || /<w:instrText\b[^>]*>[^<]*\bPAGE\b/.test(xml);
}

export function readPgNumTypeStart(xml: string): number | undefined {
  const tag = firstTag(xml, "pgNumType");
  const value = tag ? attr(tag, "start") : undefined;
  return value === undefined ? undefined : Number(value);
}

export function readJustification(xml: string): string | undefined {
  const tag = firstTag(xml, "jc");
  return tag ? attr(tag, "val") : undefined;
}
