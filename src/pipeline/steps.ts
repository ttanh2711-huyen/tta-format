import { formatBody, isFormatBodySupported, previewFormatBody } from "../office/formatBody";
import { applyFixes } from "../office/applyFix";
import { checkDocument, runFixAll } from "../office/fixAll";
import { isReadDocumentSupported, readCheckDocument } from "../office/readDocument";
import { runChecks } from "../checker/engine";
import { planFixes, type FixPlanOptions } from "../checker/fixAll";
import { isFormatOnly } from "../checker/fixes";
import { headerFooterRules } from "../checker/ruleSets";
import { applyPageNumbers, isPageNumberSupported, planPageNumberDocument } from "../office/pageNumber";
import { formatTables, surveyTables } from "../office/formatTables";
import { installStyles, isStyleInstallSupported, planStyleInstall } from "../office/installStyles";
import { applyPageSetup, checkPageSetup, isPageSetupSupported } from "../office/pageSetup";
import { summarizeApply } from "../office/pageSetupPlan";
import { summarizeInstall } from "../office/styleVerify";
import type { StyleSpec } from "../ooxml/styleSpecs";
import { toPageSetupSettings, type FormatValues } from "../settings/formatOptions";
import type { StandardProfile } from "../standards/types";
import { bodyPlan, pageNumberPlan, pagePlan, stylePlan, tablePlan } from "./planText";
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

export const pageStep: PipelineStep<FixAllContext> = {
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

export const styleStep: PipelineStep<FixAllContext> = {
  id: "styles",
  label: "2. Cài bộ style",
  order: 20,
  async plan(ctx) {
    if (!isStyleInstallSupported()) return unsupported("Word chưa hỗ trợ cài style (cần Word 2016 trở lên).");
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
    if (!isFormatBodySupported()) return unsupported("Word chưa hỗ trợ (cần Word 2016 trở lên).");
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
    if (!isFormatBodySupported()) return unsupported("Word chưa hỗ trợ (cần Word 2016 trở lên).");
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

/** Lỗi khối đầu, khối cuối của chuẩn đang chọn, đọc lại từ tài liệu. */
async function headerFooterIssues(profile: StandardProfile) {
  const rules = headerFooterRules(profile);
  return rules.length === 0 ? [] : runChecks(await readCheckDocument(), profile, rules).filter((issue) => issue.ruleId.startsWith("header.") || issue.ruleId.startsWith("footer.") || issue.ruleId.startsWith("title."));
}

export const headerFooterStep: PipelineStep<FixAllContext> = {
  id: "headerFooter",
  label: "5. Chuẩn hoá khối đầu, khối cuối",
  order: 45,
  async plan(ctx) {
    if (headerFooterRules(ctx.profile).length === 0) return unsupported(`chưa có bộ kiểm tra khối đầu, khối cuối cho chuẩn ${ctx.profile.name}.`);
    if (!isReadDocumentSupported()) return unsupported("Word chưa hỗ trợ đọc tài liệu.");
    const issues = await headerFooterIssues(ctx.profile);
    const formatFixes = issues.filter((issue) => issue.fix && isFormatOnly(issue.fix));
    const manual = issues.length - formatFixes.length;
    if (formatFixes.length === 0) {
      return { willRun: false, lines: [], skipReason: manual > 0 ? `không có lỗi định dạng tự sửa được; còn ${manual} lỗi về chữ cần bạn xem ở mục Kiểm tra văn bản.` : "khối đầu, khối cuối đã đúng định dạng." };
    }
    const lines = [`Sửa định dạng (đậm, nghiêng, cỡ chữ) ở ${formatFixes.length} chỗ trong khối đầu, khối cuối.`];
    if (manual > 0) lines.push(`Không tự sửa chữ: ${manual} lỗi còn lại để bạn xem ở mục Kiểm tra văn bản.`);
    return { willRun: true, lines };
  },
  async run(ctx) {
    const issues = await headerFooterIssues(ctx.profile);
    const fixes = issues.flatMap((issue) => (issue.fix && isFormatOnly(issue.fix) ? [issue.fix] : []));
    const result = await applyFixes(fixes);
    const manual = issues.length - fixes.length;
    const text = `Đã sửa định dạng ${result.applied} chỗ.${manual > 0 ? ` Còn ${manual} lỗi về chữ, xem ở mục Kiểm tra văn bản.` : ""}`;
    return result.skipped.length > 0 ? { status: "done", kind: "warning", text: `${text} Không sửa được: ${result.skipped.join("; ")}.` } : { status: "done", kind: manual > 0 ? "warning" : "ok", text };
  },
};

function pageNumberOptions(ctx: FixAllContext): { sizePt: number; hideFirst: boolean } {
  return {
    sizePt: typeof ctx.values.pageNumberSize === "number" ? ctx.values.pageNumberSize : 14,
    hideFirst: ctx.values.pageNumberHideFirst !== false,
  };
}

export const pageNumberStep: PipelineStep<FixAllContext> = {
  id: "pageNumber",
  label: "6. Đánh số trang",
  order: 50,
  async plan(ctx) {
    if (!isPageNumberSupported()) return unsupported("Word chưa hỗ trợ đánh số trang (cần Word 2016 trở lên).");
    const options = pageNumberOptions(ctx);
    return pageNumberPlan(await planPageNumberDocument(options), options);
  },
  async run(ctx) {
    const result = await applyPageNumbers(pageNumberOptions(ctx));
    return result.problems.length === 0
      ? { status: "done", kind: "ok", text: `Đã đánh số trang cho ${result.changed} phần.` }
      : { status: "done", kind: "warning", text: `Đã ghi nhưng đọc lại chưa đúng: ${result.problems.join("; ")}.` };
  },
};

/**
 * Bước cuối: sửa các lỗi còn sửa được tự động sau các bước trên, không đổi chữ. Giữ định dạng cục bộ thì không động
 * vào phông, cỡ chữ, căn lề, thụt dòng, giãn đoạn, giãn dòng của phần thân, vì người dùng đã chọn giữ.
 */
function remainingFixOptions(ctx: FixAllContext): FixPlanOptions {
  return {
    includeText: false,
    skipRule: ctx.localFormatting === "keep" ? (id) => id.startsWith("font.") || id.startsWith("paragraph.") : undefined,
  };
}

export const remainingFixStep: PipelineStep<FixAllContext> = {
  id: "remainingFixes",
  label: "7. Sửa các lỗi còn lại",
  order: 60,
  async plan(ctx) {
    if (!isReadDocumentSupported()) return unsupported("Word chưa hỗ trợ đọc tài liệu (cần Word 2016 trở lên).");
    const issues = await checkDocument(ctx.profile);
    const plan = planFixes(issues, remainingFixOptions(ctx));
    if (plan.fixes.length === 0) {
      return { willRun: false, lines: [], skipReason: issues.length > 0 ? `không còn lỗi sửa được tự động; còn ${issues.length} lỗi cần bạn xem ở mục Kiểm tra văn bản.` : "không còn lỗi." };
    }
    const lines = [`Sửa tự động ${plan.fixes.length} chỗ còn lệch chuẩn (không đổi chữ).`];
    if (ctx.localFormatting === "keep") lines.push("Giữ định dạng cục bộ của phần thân nên không đụng phông, cỡ chữ, căn lề, giãn đoạn, giãn dòng.");
    if (plan.manual.length > 0) lines.push(`Còn ${plan.manual.length} lỗi để bạn xem ở mục Kiểm tra văn bản.`);
    return { willRun: true, lines };
  },
  async run(ctx) {
    const result = await runFixAll(ctx.profile, remainingFixOptions(ctx));
    const left = result.remaining.length;
    const text = `Đã sửa ${result.applied} chỗ sau ${result.rounds} lượt.${left > 0 ? ` Còn ${left} lỗi hoặc cảnh báo, xem ở mục Kiểm tra văn bản.` : ""}`;
    return result.skipped.length > 0 ? { status: "done", kind: "warning", text: `${text} Không sửa được: ${result.skipped.join("; ")}.` } : { status: "done", kind: left > 0 ? "warning" : "ok", text };
  },
};

/**
 * Khung bước của "Chỉnh toàn bộ". Task sau thêm bước (số trang, khối đầu, khối cuối, sửa lỗi còn lại) bằng
 * `registry.register(...)` với `order` phù hợp; không phải sửa khung chạy hay giao diện.
 */
export function createFixAllRegistry(): StepRegistry<FixAllContext> {
  const registry = new StepRegistry<FixAllContext>();
  for (const step of [pageStep, styleStep, bodyStep, tableStep, headerFooterStep, pageNumberStep, remainingFixStep]) registry.register(step);
  return registry;
}
