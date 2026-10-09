import type { CheckDocument, Fix } from "./types";

/** Sửa chỉ định dạng, không đổi chữ: được phép chạy tự động trong "Chỉnh toàn bộ". */
export function isFormatOnly(fix: Fix): boolean {
  return fix.kind === "font";
}

/**
 * Áp một cách sửa lên bản sao của tài liệu kiểm tra: chỉ đoạn có chỉ số `fix.index` thay đổi. Dùng để xem
 * trước và để kiểm tra lại sau khi sửa mà không cần gọi Word.
 */
export function applyFix(doc: CheckDocument, fix: Fix): CheckDocument {
  return {
    ...doc,
    paragraphs: doc.paragraphs.map((p) => {
      if (p.index !== fix.index) return p;
      if (fix.kind === "replace") return { ...p, resolved: { ...p.resolved, text: p.resolved.text.replace(fix.find, fix.replacement) } };
      return {
        ...p,
        resolved: {
          ...p.resolved,
          ...(fix.bold === undefined ? {} : { bold: fix.bold }),
          ...(fix.italic === undefined ? {} : { italic: fix.italic }),
          ...(fix.sizePt === undefined ? {} : { sizePt: fix.sizePt }),
        },
      };
    }),
  };
}
