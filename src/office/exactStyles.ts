import { styleIdFromName } from "../ooxml/styleXml";
import type { StyleSpec } from "../ooxml/styleSpecs";
import type { StyleReadback } from "./styleVerify";

export interface ExactRepairPlan {
  /** Style Exactly chưa có hoặc đang sai quy tắc và chưa đoạn nào dùng (hoặc thân sắp được thay): xoá rồi cài lại bằng OOXML. */
  rebuild: StyleSpec[];
  /** Style Exactly sai quy tắc nhưng đang có đoạn dùng: xoá sẽ làm đoạn mất style, nên giữ nguyên. */
  blocked: StyleSpec[];
}

/**
 * Lập kế hoạch sửa giãn dòng Exactly cho các style. importStylesFromJson chỉ cho atLeast; chỉ cài mới bằng OOXML mới
 * đặt được Exactly (spike T3). Style đang bị sai không cập nhật tại chỗ được nên phải xoá rồi cài lại.
 */
export function planExactRepair(
  specs: readonly StyleSpec[],
  readbacks: ReadonlyMap<string, StyleReadback>,
  usedStyleIds: ReadonlySet<string>,
  bodyWillBeReplaced: boolean,
): ExactRepairPlan {
  const rebuild: StyleSpec[] = [];
  const blocked: StyleSpec[] = [];
  for (const spec of specs) {
    if (spec.paragraph.lineSpacing.kind !== "exactly") continue;
    const readback = readbacks.get(spec.name);
    if (readback?.exists && readback.lineRule === "exact") continue;
    if (!readback?.exists) rebuild.push(spec);
    else if (bodyWillBeReplaced || !usedStyleIds.has(styleIdFromName(spec.name))) rebuild.push(spec);
    else blocked.push(spec);
  }
  return { rebuild, blocked };
}
