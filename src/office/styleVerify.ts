import { lineSpacingPoints } from "../ooxml/styleJson";
import type { StyleSpec } from "../ooxml/styleSpecs";
import type { ParagraphAlignment } from "../standards/types";

/** Giá trị đọc lại từ Word cho một style (pt). `lineRule` đọc từ OOXML: auto, exact, atLeast. */
export interface StyleReadback {
  exists: boolean;
  font: { name: string; size: number; bold: boolean; italic: boolean };
  paragraph: {
    alignment: ParagraphAlignment | string;
    firstLineIndent: number;
    spaceBefore: number;
    spaceAfter: number;
    lineSpacing: number;
    keepWithNext: boolean;
  };
  lineRule?: string;
}

/** Giá trị Word lẽ ra trả về nếu style đúng như spec (dùng làm dữ liệu thử và làm mốc so sánh). */
export function readbackFromSpec(spec: StyleSpec): StyleReadback {
  return {
    exists: true,
    font: { name: spec.font.name, size: spec.font.sizePt, bold: spec.font.bold, italic: spec.font.italic },
    paragraph: {
      alignment: spec.paragraph.alignment,
      firstLineIndent: spec.paragraph.firstLineIndentPt,
      spaceBefore: spec.paragraph.spaceBeforePt,
      spaceAfter: spec.paragraph.spaceAfterPt,
      lineSpacing: lineSpacingPoints(spec.paragraph.lineSpacing),
      keepWithNext: spec.paragraph.keepWithNext,
    },
    lineRule: spec.paragraph.lineSpacing.kind === "exactly" ? "exact" : "auto",
  };
}

const POINT_TOLERANCE = 0.2;
const LINE_TOLERANCE = 0.2;

/** So style Word đọc lại với spec; trả về danh sách chênh lệch bằng tiếng Việt, rỗng nếu khớp. */
export function findStyleMismatches(spec: StyleSpec, actual: StyleReadback): string[] {
  if (!actual.exists) {
    return ["không tìm thấy style"];
  }
  const out: string[] = [];
  const num = (label: string, got: number, want: number, tolerance: number) => {
    if (!Number.isFinite(got) || Math.abs(got - want) > tolerance) {
      out.push(`${label}: đọc lại ${got}, mong đợi ${want}`);
    }
  };
  const flag = (label: string, got: boolean, want: boolean) => {
    if (got !== want) out.push(`${label}: đọc lại ${got ? "có" : "không"}, mong đợi ${want ? "có" : "không"}`);
  };

  if (actual.font.name !== spec.font.name) {
    out.push(`phông chữ: đọc lại ${actual.font.name}, mong đợi ${spec.font.name}`);
  }
  num("cỡ chữ", actual.font.size, spec.font.sizePt, 0.01);
  flag("chữ đậm", actual.font.bold, spec.font.bold);
  flag("chữ nghiêng", actual.font.italic, spec.font.italic);

  if (actual.paragraph.alignment !== spec.paragraph.alignment) {
    out.push(`căn lề: đọc lại ${actual.paragraph.alignment}, mong đợi ${spec.paragraph.alignment}`);
  }
  num("thụt dòng đầu (pt)", actual.paragraph.firstLineIndent, spec.paragraph.firstLineIndentPt, POINT_TOLERANCE);
  num("giãn trước (pt)", actual.paragraph.spaceBefore, spec.paragraph.spaceBeforePt, POINT_TOLERANCE);
  num("giãn sau (pt)", actual.paragraph.spaceAfter, spec.paragraph.spaceAfterPt, POINT_TOLERANCE);
  num("giãn dòng (pt)", actual.paragraph.lineSpacing, lineSpacingPoints(spec.paragraph.lineSpacing), LINE_TOLERANCE);
  flag("giữ với đoạn sau", actual.paragraph.keepWithNext, spec.paragraph.keepWithNext);

  if (actual.lineRule !== undefined) {
    const wanted = spec.paragraph.lineSpacing.kind === "exactly" ? "exact" : "auto";
    if (actual.lineRule !== wanted) {
      const need = wanted === "exact" ? "Exactly" : "bội số (auto)";
      out.push(`quy tắc giãn dòng: cần ${need}, Word ghi ${actual.lineRule}`);
    }
  }
  return out;
}

export interface StyleInstallItem {
  name: string;
  action: "created" | "updated";
  mismatches: string[];
}

export function summarizeInstall(items: readonly StyleInstallItem[]): { kind: "ok" | "warning"; text: string } {
  const bad = items.filter((item) => item.mismatches.length > 0);
  if (bad.length === 0) {
    const created = items.filter((i) => i.action === "created").length;
    const updated = items.length - created;
    return { kind: "ok", text: `Đã cài ${items.length} style: ${created} tạo mới, ${updated} cập nhật.` };
  }
  const details = bad.map((item) => `${item.name} (${item.mismatches.join("; ")})`).join(" ");
  return { kind: "warning", text: `Đã cài nhưng ${bad.length}/${items.length} style chưa đúng: ${details}` };
}
