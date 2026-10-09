import type { CheckDocument, Fix } from "./types";

/** Sửa chỉ định dạng, không đổi chữ: được phép chạy tự động trong "Chỉnh toàn bộ". */
export function isFormatOnly(fix: Fix): boolean {
  return fix.kind === "font";
}

/** Chỉ số đoạn mà cách sửa tác động; không có với cách sửa trang. */
export function fixIndex(fix: Fix): number | undefined {
  return fix.kind === "page" ? undefined : fix.index;
}

/**
 * Áp một cách sửa lên bản sao của tài liệu kiểm tra: chỉ đoạn có chỉ số `fix.index` (hoặc phần `fix.section`) thay đổi. Dùng để xem
 * trước và để kiểm tra lại sau khi sửa mà không cần gọi Word.
 */
export function applyFix(doc: CheckDocument, fix: Fix): CheckDocument {
  if (fix.kind === "page") {
    return {
      ...doc,
      pages: doc.pages.map((page, i) => {
        if (i !== fix.section - 1) return page;
        const paper = fix.paper ? { widthMm: 210, heightMm: 297 } : {};
        return {
          ...page,
          ...paper,
          marginsMm: { ...page.marginsMm, ...fix.marginsMm },
          headerDistanceMm: fix.headerDistanceMm ?? page.headerDistanceMm,
        };
      }),
    };
  }
  return {
    ...doc,
    paragraphs: doc.paragraphs.map((p) => {
      if (p.index !== fix.index) return p;
      const r = p.resolved;
      switch (fix.kind) {
        case "replace":
          return { ...p, resolved: { ...r, text: r.text.replace(fix.find, fix.replacement) } };
        case "font":
          return {
            ...p,
            resolved: {
              ...r,
              ...(fix.name === undefined ? {} : { fontName: fix.name }),
              ...(fix.bold === undefined ? {} : { bold: fix.bold }),
              ...(fix.italic === undefined ? {} : { italic: fix.italic }),
              ...(fix.underline === undefined ? {} : { underline: fix.underline }),
              ...(fix.sizePt === undefined ? {} : { sizePt: fix.sizePt }),
            },
          };
        case "paragraph": {
          const next = {
            ...r,
            ...(fix.align === undefined ? {} : { align: fix.align }),
            ...(fix.firstLineIndentMm === undefined ? {} : { firstLineIndentMm: fix.firstLineIndentMm }),
            ...(fix.spaceBeforePt === undefined ? {} : { spaceBeforePt: fix.spaceBeforePt }),
            ...(fix.spaceAfterPt === undefined ? {} : { spaceAfterPt: fix.spaceAfterPt }),
          };
          if (fix.lineMultiple !== undefined) return { ...p, resolved: { ...next, lineMultiple: fix.lineMultiple, lineExactPt: undefined, lineAtLeastPt: undefined } };
          return { ...p, resolved: next };
        }
        case "exact":
          return { ...p, resolved: { ...r, lineExactPt: fix.pt, lineMultiple: undefined, lineAtLeastPt: undefined } };
      }
    }),
  };
}
