/**
 * Đặt giãn dòng Exactly **trực tiếp trên đoạn** trong OOXML. API Word chỉ đặt được giãn dòng bội số
 * (kết quả spike T3), nên chuẩn Đảng (Exactly 18-22 pt) phải đi qua OOXML.
 */

/** Phần tử sau w:spacing trong w:pPr theo lược đồ; w:spacing phải đứng trước phần tử đầu tiên trong số này. */
const AFTER_SPACING = /<w:(ind|contextualSpacing|mirrorIndents|suppressOverlap|jc|textDirection|textAlignment|textboxTightWrap|outlineLvl|divId|cnfStyle|rPr|sectPr|pPrChange)\b/;

function setSpacing(pPrInner: string, attrs: string): string {
  const existing = /<w:spacing\b([^>]*?)\/>/.exec(pPrInner);
  if (existing) {
    const kept = existing[1].replace(/\s*w:(line|lineRule)="[^"]*"/g, "");
    return pPrInner.replace(existing[0], `<w:spacing${kept} ${attrs}/>`);
  }
  const element = `<w:spacing ${attrs}/>`;
  const next = AFTER_SPACING.exec(pPrInner);
  return next ? pPrInner.slice(0, next.index) + element + pPrInner.slice(next.index) : pPrInner + element;
}

function transformBody(inner: string, attrs: string): string {
  const paragraph = /<w:p(?:\s[^>]*)?\/?>/g;
  let out = "";
  let last = 0;
  let match: RegExpExecArray | null;
  while ((match = paragraph.exec(inner)) !== null) {
    const tag = match[0];
    const end = match.index + tag.length;
    out += inner.slice(last, match.index);
    if (tag.endsWith("/>")) {
      out += `${tag.slice(0, -2)}><w:pPr><w:spacing ${attrs}/></w:pPr></w:p>`;
      last = end;
      continue;
    }
    out += tag;
    const rest = inner.slice(end);
    const pPr = /^<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(rest);
    if (pPr) {
      out += `<w:pPr>${setSpacing(pPr[1], attrs)}</w:pPr>`;
      last = end + pPr[0].length;
    } else if (rest.startsWith("<w:pPr/>")) {
      out += `<w:pPr><w:spacing ${attrs}/></w:pPr>`;
      last = end + "<w:pPr/>".length;
    } else {
      out += `<w:pPr><w:spacing ${attrs}/></w:pPr>`;
      last = end;
    }
  }
  return out + inner.slice(last);
}

/**
 * Gói OOXML (flat OPC) với mọi đoạn trong phần thân được đặt giãn dòng Exactly `pt`. Giữ nguyên giãn
 * trước, sau và mọi thuộc tính khác; không đụng phần styles. Chạy lại cho cùng kết quả.
 */
export function withExactLineSpacing(ooxml: string, pt: number): string {
  const body = /<w:body>([\s\S]*)<\/w:body>/.exec(ooxml);
  if (!body) {
    return ooxml;
  }
  const attrs = `w:line="${Math.round(pt * 20)}" w:lineRule="exact"`;
  const innerStart = body.index + "<w:body>".length;
  return ooxml.slice(0, innerStart) + transformBody(body[1], attrs) + ooxml.slice(innerStart + body[1].length);
}

/**
 * Bỏ w:sectPr ở cấp thân (thuộc tính trang của cả tài liệu) khỏi gói OOXML; giữ w:sectPr nằm trong đoạn
 * (ngắt phần thật). Gói lấy từ `getOoxml()` của một đoạn vẫn kèm thuộc tính trang, chèn lại sẽ ghi đè
 * thiết lập trang hiện có.
 */
export function stripBodySectPr(ooxml: string): string {
  const body = /<w:body>([\s\S]*)<\/w:body>/.exec(ooxml);
  if (!body) {
    return ooxml;
  }
  const inner = body[1];
  const start = inner.lastIndexOf("<w:sectPr");
  if (start < 0) {
    return ooxml;
  }
  const tail = inner.slice(start);
  if (!/^<w:sectPr\b(?:[^>]*\/>|[\s\S]*<\/w:sectPr>)\s*$/.test(tail)) {
    return ooxml;
  }
  const head = inner.slice(0, start);
  if (head.lastIndexOf("<w:pPr>") > head.lastIndexOf("</w:pPr>")) {
    return ooxml;
  }
  const innerStart = body.index + "<w:body>".length;
  return ooxml.slice(0, innerStart) + head + ooxml.slice(innerStart + inner.length);
}

/** OOXML của một đoạn sẵn sàng chèn lại: bỏ thuộc tính trang cấp thân rồi đặt giãn dòng Exactly. */
export function prepareParagraphOoxml(ooxml: string, pt: number): string {
  return withExactLineSpacing(stripBodySectPr(ooxml), pt);
}
