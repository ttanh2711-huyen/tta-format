import { planTemplate, type TemplatePlan } from "../checker/templatePlan";
import type { StandardProfile } from "../standards/types";
import { readCheckDocument } from "./readDocument";

/** Tiền tố tag của các Content Control do add-in tạo, để nhận ra và điền lại sau này. */
export const TEMPLATE_TAG_PREFIX = "tta.";
/** Màu viền của Content Control (vàng đậm, dễ thấy trên nền trắng). */
const CONTROL_COLOR = "#E0A800";

export interface MakeTemplateResult {
  plan: TemplatePlan;
  applied: number;
}

/** Chỉ đọc: các vùng sẽ thay bằng Content Control. */
export async function previewTemplate(profile: StandardProfile): Promise<TemplatePlan> {
  return planTemplate(await readCheckDocument(), profile);
}

/**
 * Biến văn bản đang mở thành mẫu (issue #2): mỗi vùng thay đổi (số, ngày tháng, trích yếu, căn cứ, thân, nơi nhận,
 * người ký…) được thay bằng chữ gợi ý và bọc trong Content Control có tiêu đề, giữ nguyên bố cục trang, bảng thể thức
 * và style. Làm từ vùng cuối lên đầu để chỉ số đoạn phía trên không lệch khi gộp nhiều đoạn thành một.
 */
export async function makeTemplate(profile: StandardProfile): Promise<MakeTemplateResult> {
  const plan = planTemplate(await readCheckDocument(), profile);
  if (plan.fields.length === 0) return { plan, applied: 0 };
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();
    let applied = 0;
    for (const f of [...plan.fields].sort((a, b) => b.start - a.start)) {
      const first = paragraphs.items[f.start];
      const last = paragraphs.items[f.end];
      if (!first || !last) continue;
      let range = f.start === f.end ? first.getRange("Content") : first.getRange("Content").expandTo(last.getRange("Content"));
      if (f.placeholder !== null) range = range.insertText(f.placeholder, "Replace");
      const control = range.insertContentControl();
      control.tag = `${TEMPLATE_TAG_PREFIX}${f.tag}`;
      control.title = f.title;
      control.appearance = "BoundingBox";
      control.color = CONTROL_COLOR;
      applied += 1;
    }
    await context.sync();
    return { plan, applied };
  });
}
