import { formatBody, isFormatBodySupported, previewFormatBody } from "../office/formatBody";
import { formatTables, surveyTables } from "../office/formatTables";
import { installStyles, isStyleInstallSupported, planStyleInstall } from "../office/installStyles";
import { applyPageSetup, checkPageSetup, isPageSetupSupported } from "../office/pageSetup";
import { summarizeApply } from "../office/pageSetupPlan";
import { summarizeInstall } from "../office/styleVerify";
import type { StyleSpec } from "../ooxml/styleSpecs";
import { toPageSetupSettings, type FormatValues } from "../settings/formatOptions";
import type { StandardProfile } from "../standards/types";
import { bodyPlan, pagePlan, stylePlan, tablePlan } from "./planText";
import { StepRegistry, type PipelineStep } from "./runner";

/** Mọi thứ các bước cần biết, lấy từ chuẩn đang chọn và Tuỳ chọn nhanh tại lúc bấm nút. */
export interface FixAllContext {
  profile: StandardProfile;
  values: FormatValues;
  specs: StyleSpec[];
  localFormatting: "keep" | "clear";
}

function vn(n: number): string {
  return String(n).replace(".", ",");
}

function unsupported(reason: string) {
  return { willRun: false, lines: [], skipReason: reason };
}

const pageStep: PipelineStep<FixAllContext> = {
  id: "page",
  label: "1. Thiết lập trang",
  order: 10,
  async plan(ctx) {
    if (!isPageSetupSupported()) return unsupported("Word chưa hỗ trợ thiết lập trang (cần Microsoft 365 bản 2507 trở lên).");
    const settings = toPageSetupSettings(ctx.values);
    const { sections, mismatches } = await checkPageSetup(settings);
    const { top, bottom, left, right } = settings.marginsMm;
    return pagePlan({ sections, mismatches, marginsText: `${vn(top)}/${vn(bottom)}/${vn(left)}/${vn(right)} mm` });
  },
  async run(ctx) {
    const settings = toPageSetupSettings(ctx.values);
    const result = await applyPageSetup(settings, { defaultFontName: ctx.profile.font.name });
    const summary = summarizeApply(settings, result.sections, result.mismatches);
    return { status: "done", kind: summary.kind, text: [summary.text, result.fontNote].filter(Boolean).join(" ") };
  },
};

const styleStep: PipelineStep<FixAllContext> = {
  id: "styles",
  label: "2. Cài bộ style",
  order: 20,
  async plan(ctx) {
    if (!isStyleInstallSupported()) return unsupported("Word chưa hỗ trợ cài style (cần Microsoft 365 bản 2308 trở lên).");
    return stylePlan(await planStyleInstall(ctx.specs), ctx.profile.styles.prefix);
  },
  async run(ctx) {
    const result = await installStyles(ctx.specs);
    const summary = summarizeInstall(result.items);
    return { status: "done", kind: summary.kind, text: [summary.text, ...result.notes].join(" ") };
  },
};

function exactPtOf(ctx: FixAllContext): number | undefined {
  const content = ctx.specs.find((spec) => spec.name === `${ctx.profile.styles.prefix}NoiDung`);
  return content?.paragraph.lineSpacing.kind === "exactly" ? content.paragraph.lineSpacing.pt : undefined;
}

const bodyStep: PipelineStep<FixAllContext> = {
  id: "body",
  label: "3. Định dạng phần thân",
  order: 30,
  async plan(ctx) {
    if (!isFormatBodySupported()) return unsupported("Word chưa hỗ trợ (cần Microsoft 365 bản 2308 trở lên).");
    return bodyPlan(await previewFormatBody(ctx.profile, "document"), exactPtOf(ctx), `${ctx.profile.styles.prefix}NoiDung`);
  },
  async run(ctx) {
    const result = await formatBody(ctx.profile, { scope: "document", localFormatting: ctx.localFormatting, specs: ctx.specs });
    return { status: result.applied > 0 ? "done" : "skipped", kind: result.summary.kind, text: result.summary.text };
  },
};

const tableStep: PipelineStep<FixAllContext> = {
  id: "tables",
  label: "4. Định dạng bảng biểu",
  order: 40,
  async plan(ctx) {
    if (!isFormatBodySupported()) return unsupported("Word chưa hỗ trợ (cần Microsoft 365 bản 2308 trở lên).");
    return tablePlan(await surveyTables(ctx.profile), `${ctx.profile.styles.prefix}BangDuLieu`);
  },
  async run(ctx) {
    const result = await formatTables(ctx.profile, ctx.localFormatting);
    return {
      status: "done",
      kind: "ok",
      text: `Đã áp style chữ bảng cho ${result.paragraphsToStyle} đoạn trong ${result.data} bảng có viền; giữ nguyên ${result.layout} bảng không viền.`,
    };
  },
};

/**
 * Khung bước của "Chỉnh toàn bộ". Task sau thêm bước (số trang, khối đầu, khối cuối, sửa lỗi còn lại) bằng
 * `registry.register(...)` với `order` phù hợp; không phải sửa khung chạy hay giao diện.
 */
export function createFixAllRegistry(): StepRegistry<FixAllContext> {
  const registry = new StepRegistry<FixAllContext>();
  for (const step of [pageStep, styleStep, bodyStep, tableStep]) registry.register(step);
  return registry;
}
