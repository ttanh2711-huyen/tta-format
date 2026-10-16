import { buildStyleSpecs } from "../ooxml/styleSpecs";
import type { FormatValues } from "../settings/formatOptions";
import { toPageSetupSettings, toStyleOptions } from "../settings/formatOptions";
import type { StandardProfile } from "../standards/types";
import { installStyles, isStyleInstallSupported } from "./installStyles";
import { applyPageSetup, isPageSetupSupported } from "./pageSetup";
import { summarizeApply } from "./pageSetupPlan";
import { summarizeInstall } from "./styleVerify";

export interface ApplyOutcome {
  kind: "ok" | "warning" | "error";
  text: string;
}

/** Áp lề, khổ giấy, khoảng cách header/footer, lề đối xứng theo tuỳ chọn hiện hành lên mọi phần. */
export async function applyPageOptions(values: FormatValues): Promise<ApplyOutcome> {
  if (!isPageSetupSupported()) {
    return { kind: "error", text: "Thiết lập trang bằng add-in cần Microsoft 365 bản 2507 trở lên (WordApiDesktop 1.3); bản Word khác hãy đặt trong Layout > Margins hoặc dùng bộ macro VBA." };
  }
  const settings = toPageSetupSettings(values);
  const result = await applyPageSetup(settings);
  return summarizeApply(settings, result.sections, result.mismatches);
}

async function contentStyleExists(name: string): Promise<boolean> {
  return Word.run(async (context) => {
    const style = context.document.getStyles().getByNameOrNullObject(name);
    style.load("isNullObject");
    await context.sync();
    return !style.isNullObject;
  });
}

/**
 * Cập nhật bộ style theo tuỳ chọn hiện hành để đoạn đã dùng style đổi ngay. Chưa cài bộ style thì
 * không cài thay người dùng: chỉ báo giá trị sẽ có hiệu lực khi bấm "Cài bộ style".
 */
export async function applyStyleOptions(profile: StandardProfile, values: FormatValues): Promise<ApplyOutcome> {
  if (!isStyleInstallSupported()) {
    return { kind: "error", text: "Cập nhật style cần Word 2016 trở lên." };
  }
  if (!(await contentStyleExists(`${profile.styles.prefix}NoiDung`))) {
    return { kind: "ok", text: "Đã ghi nhận. Tài liệu chưa có bộ style, giá trị sẽ có hiệu lực khi bấm \"Cài bộ style\"." };
  }
  const result = await installStyles(buildStyleSpecs(profile, toStyleOptions(profile, values)));
  const summary = summarizeInstall(result.items);
  const note =
    profile.content.lineSpacing.kind === "exactly"
      ? " Đoạn đã đặt Exactly giữ giá trị cũ cho đến khi bấm lại \"Định dạng nội dung\"."
      : "";
  return { kind: summary.kind, text: [summary.text, ...result.notes].join(" ") + note };
}
