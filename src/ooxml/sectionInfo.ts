import type { PageInfo } from "../checker/types";
import { bodyPart } from "./facts";

/**
 * Thiết lập trang đọc từ OOXML của thân văn bản (các thẻ `w:sectPr`), dùng khi Word không có API thiết lập trang
 * (WordApiDesktop 1.3 chỉ có trên Microsoft 365): Word 2016, 2019, 2021 vẫn kiểm tra được khổ giấy, lề. Lề đối xứng nằm
 * trong phần settings của tài liệu nên ở đây coi là không có.
 */

const TWIPS_PER_MM = 1440 / 25.4;
const mm = (twips: number): number => Math.round((twips / TWIPS_PER_MM) * 100) / 100;

function num(tag: string | undefined, name: string): number | undefined {
  if (!tag) return undefined;
  const value = new RegExp(`\\bw:${name}="(-?\\d+)"`).exec(tag)?.[1];
  return value === undefined ? undefined : Number(value);
}

/** Thiết lập trang của từng phần, theo thứ tự trong tài liệu; rỗng nếu OOXML không có `w:sectPr`. */
export function pagesFromOoxml(xml: string): PageInfo[] {
  const body = bodyPart(xml) || xml;
  const sections = body.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? [];
  return sections.map((sect) => {
    const size = /<w:pgSz\b[^>]*>/.exec(sect)?.[0];
    const margin = /<w:pgMar\b[^>]*>/.exec(sect)?.[0];
    return {
      widthMm: mm(num(size, "w") ?? 11906),
      heightMm: mm(num(size, "h") ?? 16838),
      marginsMm: { top: mm(Math.abs(num(margin, "top") ?? 1440)), bottom: mm(Math.abs(num(margin, "bottom") ?? 1440)), left: mm(num(margin, "left") ?? 1800), right: mm(num(margin, "right") ?? 1800) },
      headerDistanceMm: mm(num(margin, "header") ?? 720),
      footerDistanceMm: mm(num(margin, "footer") ?? 720),
      mirrorMargins: false,
    };
  });
}
