import type { StyleInstallPlan } from "../office/installStyles";
import type { FormatBodyPreview } from "../office/formatBody";
import type { TableSurvey } from "../office/formatTables";
import { describeSkipped } from "../office/paragraphClassifier";
import { isPlanEmpty, type PageNumberOptions, type SectionPlan } from "../ooxml/pageNumberHeader";
import type { StepPlan } from "./runner";

/** Câu chữ của hộp xem trước cho từng bước, tách khỏi mã gọi Word để test được. */

const MAX_MISMATCHES = 5;

export function pagePlan(input: { sections: number; mismatches: readonly string[]; marginsText: string }): StepPlan {
  if (input.mismatches.length === 0) {
    return { willRun: false, lines: [], skipReason: `trang đã đúng thiết lập (A4, lề ${input.marginsText}).` };
  }
  const shown = input.mismatches.slice(0, MAX_MISMATCHES);
  const lines = [`Đặt khổ A4 dọc, lề ${input.marginsText} cho ${input.sections} phần.`, `Đang lệch: ${shown.join("; ")}`];
  const more = input.mismatches.length - shown.length;
  if (more > 0) lines.push(`… và ${more} điểm lệch khác.`);
  return { willRun: true, lines };
}

export function stylePlan(plan: StyleInstallPlan, prefix: string): StepPlan {
  if (plan.toCreate === 0 && plan.toUpdate === 0) {
    return { willRun: false, lines: [], skipReason: `bộ style ${prefix}* đã đúng.` };
  }
  return { willRun: true, lines: [`Cài bộ style ${prefix}*: tạo ${plan.toCreate} style, cập nhật ${plan.toUpdate} style.`] };
}

export function bodyPlan(preview: FormatBodyPreview, exactPt: number | undefined, styleName: string): StepPlan {
  const skipText = describeSkipped(preview.skipped);
  if (preview.targets === 0) {
    return { willRun: false, lines: [], skipReason: ["không có đoạn nào cần định dạng.", skipText].filter(Boolean).join(" ") };
  }
  const lines = [`Áp style ${styleName} cho ${preview.targets} đoạn.`];
  if (skipText) lines.push(skipText);
  if (exactPt !== undefined) {
    lines.push(`Giãn dòng Exactly ${exactPt} pt đặt trực tiếp trên từng đoạn bằng cách ghi lại đoạn; bình luận hoặc đánh dấu gắn trong đoạn có thể không còn.`);
  }
  return { willRun: true, lines };
}

function keptTables(survey: TableSurvey): string {
  const parts: string[] = [];
  if (survey.layout > 0) parts.push(`${survey.layout} bảng không viền`);
  if (survey.nested > 0) parts.push(`${survey.nested} bảng lồng nhau`);
  return parts.join(" và ");
}

export function tablePlan(survey: TableSurvey, styleName: string): StepPlan {
  const kept = keptTables(survey);
  if (survey.data === 0) {
    const note = survey.layout > 0 ? ` Giữ nguyên ${survey.layout} bảng không viền (đầu văn bản, nơi nhận, chữ ký).` : "";
    return { willRun: false, lines: [], skipReason: `không có bảng dữ liệu cần định dạng.${note}` };
  }
  if (survey.paragraphsToStyle === 0) {
    return { willRun: false, lines: [], skipReason: "bảng dữ liệu đã dùng style chữ bảng." };
  }
  const lines = [`Áp style ${styleName} cho ${survey.paragraphsToStyle} đoạn trong ${survey.data} bảng có viền.`];
  if (kept) lines.push(`Giữ nguyên ${kept}.`);
  return { willRun: true, lines };
}

export function pageNumberPlan(plan: readonly SectionPlan[], options: PageNumberOptions): StepPlan {
  if (isPlanEmpty(plan)) {
    const first = options.hideFirst ? ", trang 1 không hiện số" : "";
    return { willRun: false, lines: [], skipReason: `số trang đã đúng (giữa trang, ${options.sizePt} pt${first}).` };
  }
  const lines: string[] = [];
  const write = plan.filter((item) => item.writePrimary).length;
  if (write > 0) lines.push(`Ghi số trang (căn giữa, ${options.sizePt} pt) vào header của ${write} phần.`);
  const toggled = plan.filter((item) => item.setDifferentFirst !== undefined).length;
  if (toggled > 0) lines.push(`Đặt lại "trang đầu khác biệt" cho ${toggled} phần.`);
  if (plan.some((item) => item.clearFirst)) lines.push("Xoá số trang đang nằm ở header trang đầu (trang 1 không hiện số).");
  return { willRun: true, lines };
}
