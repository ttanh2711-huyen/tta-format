import type { PageSetupSettings } from "../office/pageSetupPlan";
import { defaultStyleOptions } from "../ooxml/styleSpecs";
import { isFirstLineIndentValid, isFontSizeValid, isLineSpacingValid, isMarginValid } from "../standards/validate";
import type { MarginSide, StandardProfile, StyleOptions } from "../standards/types";

/**
 * Registry các tuỳ chọn định dạng nhanh (T6A). Mỗi chuẩn dựng registry riêng từ chính `StandardProfile`
 * (khoảng cho phép, mặc định), nên thêm tuỳ chọn chỉ cần thêm một mục ở `buildRegistry`; giao diện
 * sinh từ registry. Mã ở đây không rẽ nhánh theo mã chuẩn.
 */

export type OptionValue = number | string | boolean;
export type FormatValues = Record<string, OptionValue>;

export type GroupId = "page" | "font" | "paragraph" | "header" | "footer" | "table" | "pageNumber" | "behavior";

/** now: đổi tài liệu ngay; onCreate: dùng khi chèn khối hoặc tạo văn bản (T7 trở đi); behavior: cách chạy các nút. */
export type Effect = "now" | "onCreate" | "behavior";

export interface Choice {
  value: OptionValue;
  label: string;
  /** Chọn giá trị này là ngoại lệ so với chuẩn; nêu lý do để hỏi xác nhận. */
  exception?: string;
}

export interface OptionDef {
  id: string;
  group: GroupId;
  label: string;
  kind: "number" | "boolean" | "choice";
  /** Đơn vị hiển thị cạnh ô nhập: pt, mm, dòng… */
  unit?: string;
  effect: Effect;
  /** Giá trị chọn nhanh. */
  choices: Choice[];
  default: OptionValue;
  /** Cho nhập số tuỳ ý ("Tuỳ chỉnh…"). */
  custom: boolean;
  /** Trả về lý do khi giá trị nằm ngoài quy định của chuẩn; null nếu đúng quy định. */
  rule?: (value: number) => string | null;
  /** Giới hạn cứng: ngoài khoảng này là sai hẳn, không cho xác nhận ngoại lệ. */
  hard?: { min: number; max: number };
}

export interface GroupDef {
  id: GroupId;
  label: string;
}

export interface FormatRegistry {
  standardName: string;
  groups: GroupDef[];
  options: OptionDef[];
}

export interface FormatState {
  values: FormatValues;
  /** Mã các tuỳ chọn đang giữ giá trị ngoài quy định do người dùng xác nhận (bộ kiểm tra không báo lại). */
  exceptions: string[];
}

const GROUPS: GroupDef[] = [
  { id: "page", label: "Trang" },
  { id: "font", label: "Chữ" },
  { id: "paragraph", label: "Đoạn" },
  { id: "header", label: "Đầu văn bản" },
  { id: "footer", label: "Cuối văn bản" },
  { id: "table", label: "Bảng biểu" },
  { id: "pageNumber", label: "Số trang" },
  { id: "behavior", label: "Hành vi" },
];

const SIDES: ReadonlyArray<{ side: MarginSide; id: string; label: string; name: string }> = [
  { side: "top", id: "marginTop", label: "Lề trên", name: "trên" },
  { side: "bottom", id: "marginBottom", label: "Lề dưới", name: "dưới" },
  { side: "left", id: "marginLeft", label: "Lề trái", name: "trái" },
  { side: "right", id: "marginRight", label: "Lề phải", name: "phải" },
];

function vn(n: number): string {
  return String(Math.round(n * 1000) / 1000).replace(".", ",");
}

function integers(min: number, max: number): number[] {
  const out: number[] = [];
  for (let n = Math.ceil(min); n <= Math.floor(max); n += 1) {
    out.push(n);
  }
  return out;
}

function unique(values: number[]): number[] {
  return [...new Set(values)].sort((a, b) => a - b);
}

function numberChoices(values: number[], unit: string): Choice[] {
  return values.map((value) => ({ value, label: `${vn(value)} ${unit}` }));
}

function rangeText(min: number, max: number, unit: string): string {
  return min === max ? `quy định đúng ${vn(min)} ${unit}` : `chỉ cho phép trong khoảng ${vn(min)}–${vn(max)} ${unit}`;
}

function inRange(profile: StandardProfile, label: string, unit: string, min: number, max: number): (value: number) => string | null {
  return (value) => (value >= min && value <= max ? null : `${label} ${vn(value)} ${unit} ngoài quy định: chuẩn ${profile.name} ${rangeText(min, max, unit)}.`);
}

export function buildRegistry(profile: StandardProfile): FormatRegistry {
  const options: OptionDef[] = [];
  const { pageDefaults, styles, content, pageNumber } = profile;
  const defaults = styles.defaults;

  for (const { side, id, label } of SIDES) {
    const range = profile.page.marginsMm[side];
    const fallback = pageDefaults.marginsMm[side];
    options.push({
      id,
      group: "page",
      label,
      kind: "number",
      unit: "mm",
      effect: "now",
      choices: numberChoices(unique([range.min, fallback, range.max]), "mm"),
      default: fallback,
      custom: true,
      hard: { min: 0, max: 100 },
      rule: (value) =>
        isMarginValid(profile, side, value) ? null : `${label} ${vn(value)} mm ngoài quy định: chuẩn ${profile.name} ${rangeText(range.min, range.max, "mm")}.`,
    });
  }
  for (const [id, label, value] of [
    ["headerDistance", "Khoảng cách header", pageDefaults.headerDistanceMm],
    ["footerDistance", "Khoảng cách footer", pageDefaults.footerDistanceMm],
  ] as const) {
    options.push({ id, group: "page", label, kind: "number", unit: "mm", effect: "now", choices: numberChoices([value], "mm"), default: value, custom: true, hard: { min: 0, max: 50 } });
  }
  if (profile.page.duplexBackMarginsMm) {
    options.push({ id: "duplex", group: "page", label: "In hai mặt (lề đối xứng)", kind: "boolean", effect: "now", choices: [], default: false, custom: false });
  }

  const body = profile.elements.noiDung;
  options.push({
    id: "fontSize",
    group: "font",
    label: "Cỡ chữ nội dung",
    kind: "number",
    unit: "pt",
    effect: "now",
    choices: numberChoices(integers(body.sizePt.min, body.sizePt.max), "pt"),
    default: defaults.fontSizePt,
    custom: true,
    hard: { min: 8, max: 30 },
    rule: (value) => (isFontSizeValid(profile, "noiDung", value) ? null : `Cỡ chữ ${vn(value)} pt ngoài quy định: chuẩn ${profile.name} ${rangeText(body.sizePt.min, body.sizePt.max, "pt")}.`),
  });

  options.push({
    id: "alignment",
    group: "paragraph",
    label: "Căn lề nội dung",
    kind: "choice",
    effect: "now",
    choices: [
      { value: "Justified", label: "Đều hai bên" },
      { value: "Left", label: "Trái", exception: `Chuẩn ${profile.name} quy định nội dung dàn đều hai bên.` },
    ],
    default: "Justified",
    custom: false,
  });
  options.push({
    id: "firstLineIndent",
    group: "paragraph",
    label: "Thụt dòng đầu",
    kind: "number",
    unit: "mm",
    effect: "now",
    choices: content.firstLineIndentMm.map((value) => ({ value, label: `${vn(value / 10)} cm` })),
    default: defaults.firstLineIndentMm,
    custom: true,
    hard: { min: 0, max: 50 },
    rule: (value) =>
      isFirstLineIndentValid(profile, value) ? null : `Thụt dòng đầu ${vn(value)} mm ngoài quy định: chuẩn ${profile.name} chỉ cho phép ${content.firstLineIndentMm.map((mm) => `${vn(mm)} mm`).join(" hoặc ")}.`,
  });
  options.push({
    id: "spacing",
    group: "paragraph",
    label: "Giãn đoạn trước và sau",
    kind: "number",
    unit: "pt",
    effect: "now",
    choices: numberChoices(unique([content.paragraphSpacingMinPt, defaults.spaceBeforePt, 8, 12].filter((n) => n >= content.paragraphSpacingMinPt)), "pt"),
    default: defaults.spaceBeforePt,
    custom: true,
    hard: { min: 0, max: 72 },
    rule: (value) => (value >= content.paragraphSpacingMinPt ? null : `Giãn đoạn ${vn(value)} pt ngoài quy định: chuẩn ${profile.name} yêu cầu không dưới ${vn(content.paragraphSpacingMinPt)} pt.`),
  });

  const rule = content.lineSpacing;
  if (rule.kind === "multiple") {
    const current = defaults.lineSpacing.kind === "multiple" ? defaults.lineSpacing.lines : rule.min;
    options.push({
      id: "lineSpacing",
      group: "paragraph",
      label: "Giãn dòng",
      kind: "number",
      unit: "dòng",
      effect: "now",
      choices: unique([rule.min, current, rule.max]).map((value) => ({ value, label: vn(value) })),
      default: current,
      custom: true,
      hard: { min: 0.5, max: 3 },
      rule: (value) => (isLineSpacingValid(profile, { kind: "multiple", lines: value }) ? null : `Giãn dòng ${vn(value)} ngoài quy định: chuẩn ${profile.name} ${rangeText(rule.min, rule.max, "dòng")}.`),
    });
  } else {
    const current = defaults.lineSpacing.kind === "exactly" ? defaults.lineSpacing.pt : rule.minPt;
    options.push({
      id: "lineSpacing",
      group: "paragraph",
      label: "Giãn dòng (Exactly)",
      kind: "number",
      unit: "pt",
      effect: "now",
      choices: numberChoices(integers(rule.minPt, rule.maxPt), "pt"),
      default: current,
      custom: true,
      hard: { min: 6, max: 72 },
      rule: (value) => (isLineSpacingValid(profile, { kind: "exactly", pt: value }) ? null : `Giãn dòng ${vn(value)} pt ngoài quy định: chuẩn ${profile.name} ${rangeText(rule.minPt, rule.maxPt, "pt")}.`),
    });
  }

  const sizeOptions = (group: "header" | "footer"): OptionDef[] =>
    styles.blockSizes
      .filter((block) => block.group === group)
      .map((block) => {
        const element = profile.elements[block.elementId];
        return {
          id: `size.${block.elementId}`,
          group,
          label: `Cỡ ${element.label}`,
          kind: "number" as const,
          unit: "pt",
          effect: "onCreate" as const,
          choices: numberChoices(integers(element.sizePt.min, element.sizePt.max), "pt"),
          default: block.defaultPt,
          custom: true,
          hard: { min: 8, max: 30 },
          rule: inRange(profile, `Cỡ ${element.label}`, "pt", element.sizePt.min, element.sizePt.max),
        };
      });
  options.push(...sizeOptions("header"));
  const ratioLabel = (ratio: number) => (Math.abs(ratio - 1 / 3) < 1e-6 ? "1/3" : Math.abs(ratio - 1 / 2) < 1e-6 ? "1/2" : vn(Math.round(ratio * 100) / 100));
  for (const [id, range, what] of [
    ["agencyLineRatio", profile.agencyLineRatio, "tên cơ quan"],
    ["titleLineRatio", profile.titleLineRatio, "trích yếu"],
  ] as const) {
    if (!range) continue;
    const { min, max } = range;
    options.push({
      id,
      group: "header",
      label: `Độ dài đường kẻ dưới ${what} (so với dòng chữ)`,
      kind: "number",
      effect: "onCreate",
      choices: unique([min, max]).map((value) => ({ value, label: ratioLabel(value) })),
      default: max,
      custom: true,
      hard: { min: 0.1, max: 1 },
      rule: (value) => (value >= min - 1e-6 && value <= max + 1e-6 ? null : `Đường kẻ dài ${vn(Math.round(value * 100) / 100)} dòng chữ ngoài quy định: chuẩn ${profile.name} cho phép từ ${ratioLabel(min)} đến ${ratioLabel(max)}.`),
    });
  }
  options.push(...sizeOptions("footer"));
  options.push({ id: "signGap", group: "footer", label: "Số dòng trống chỗ ký", kind: "number", unit: "dòng", effect: "onCreate", choices: numberChoices([4, 5, 6, 7], "dòng"), default: 6, custom: true, hard: { min: 1, max: 12 } });

  options.push({ id: "tableFontSize", group: "table", label: "Cỡ chữ trong bảng", kind: "number", unit: "pt", effect: "now", choices: numberChoices([13, 14], "pt"), default: styles.table.fontSizePt, custom: true, hard: { min: 8, max: 20 } });
  options.push({ id: "tableSpacing", group: "table", label: "Giãn ô trước và sau", kind: "number", unit: "pt", effect: "now", choices: numberChoices([0, 3], "pt"), default: styles.table.spaceBeforePt, custom: true, hard: { min: 0, max: 24 } });
  options.push({ id: "tableBorder", group: "table", label: "Viền bảng", kind: "boolean", effect: "onCreate", choices: [], default: true, custom: false });
  options.push({ id: "tableHeaderBold", group: "table", label: "Hàng tiêu đề đậm, căn giữa", kind: "boolean", effect: "onCreate", choices: [], default: true, custom: false });

  options.push({
    id: "pageNumberSize",
    group: "pageNumber",
    label: "Cỡ số trang",
    kind: "number",
    unit: "pt",
    effect: "onCreate",
    choices: numberChoices(integers(pageNumber.sizePt.min, pageNumber.sizePt.max), "pt"),
    default: pageNumber.sizePt.max,
    custom: true,
    hard: { min: 8, max: 30 },
    rule: inRange(profile, "Cỡ số trang", "pt", pageNumber.sizePt.min, pageNumber.sizePt.max),
  });
  options.push({ id: "pageNumberHideFirst", group: "pageNumber", label: "Trang 1 không hiện số", kind: "boolean", effect: "onCreate", choices: [], default: pageNumber.hideOnFirstPage, custom: false });
  options.push({ id: "pageNumberAppendixSeparate", group: "pageNumber", label: "Phụ lục đánh số riêng", kind: "boolean", effect: "onCreate", choices: [], default: pageNumber.appendixNumberedSeparately, custom: false });

  options.push({
    id: "localFormatting",
    group: "behavior",
    label: "Định dạng cục bộ (đậm, nghiêng…)",
    kind: "choice",
    effect: "behavior",
    choices: [
      { value: "keep", label: "Giữ lại" },
      { value: "clear", label: "Xoá để theo style" },
    ],
    default: "keep",
    custom: false,
  });

  return { standardName: profile.name, groups: GROUPS, options };
}

export function defaultValues(registry: FormatRegistry): FormatValues {
  return Object.fromEntries(registry.options.map((o) => [o.id, o.default]));
}

export function initialState(registry: FormatRegistry): FormatState {
  return { values: defaultValues(registry), exceptions: [] };
}

export type ValueCheck = { status: "ok" } | { status: "exception"; reason: string } | { status: "invalid"; reason: string };

export function checkValue(def: OptionDef, value: unknown): ValueCheck {
  if (def.kind === "boolean") {
    return typeof value === "boolean" ? { status: "ok" } : { status: "invalid", reason: `"${def.label}" chỉ nhận bật hoặc tắt.` };
  }
  if (def.kind === "choice") {
    const choice = def.choices.find((c) => c.value === value);
    if (!choice) {
      return { status: "invalid", reason: `"${def.label}": giá trị không nằm trong các lựa chọn.` };
    }
    return choice.exception ? { status: "exception", reason: choice.exception } : { status: "ok" };
  }
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return { status: "invalid", reason: `"${def.label}": hãy nhập một số hợp lệ.` };
  }
  if (def.hard && (value < def.hard.min || value > def.hard.max)) {
    return { status: "invalid", reason: `"${def.label}": giá trị phải từ ${vn(def.hard.min)} đến ${vn(def.hard.max)} ${def.unit ?? ""}`.trim() + "." };
  }
  if (!def.custom && !def.choices.some((c) => c.value === value)) {
    return { status: "invalid", reason: `"${def.label}": giá trị không nằm trong các lựa chọn.` };
  }
  const reason = def.rule?.(value) ?? null;
  return reason === null ? { status: "ok" } : { status: "exception", reason };
}

export type SetResult =
  | { status: "applied"; state: FormatState }
  | { status: "needsConfirm"; reason: string }
  | { status: "rejected"; reason: string };

function withException(exceptions: readonly string[], id: string, on: boolean): string[] {
  const rest = exceptions.filter((e) => e !== id);
  return on ? [...rest, id] : rest;
}

export function setOption(registry: FormatRegistry, state: FormatState, id: string, value: unknown, options: { confirmed?: boolean } = {}): SetResult {
  const def = registry.options.find((o) => o.id === id);
  if (!def) {
    return { status: "rejected", reason: `Không có tuỳ chọn "${id}" trong chuẩn ${registry.standardName}.` };
  }
  const check = checkValue(def, value);
  if (check.status === "invalid") {
    return { status: "rejected", reason: check.reason };
  }
  if (check.status === "exception" && options.confirmed !== true) {
    return { status: "needsConfirm", reason: check.reason };
  }
  return {
    status: "applied",
    state: { values: { ...state.values, [id]: value as OptionValue }, exceptions: withException(state.exceptions, id, check.status === "exception") },
  };
}

export function resetOption(registry: FormatRegistry, state: FormatState, id: string): FormatState {
  const def = registry.options.find((o) => o.id === id);
  if (!def) {
    return state;
  }
  return { values: { ...state.values, [id]: def.default }, exceptions: withException(state.exceptions, id, false) };
}

export function resetGroup(registry: FormatRegistry, state: FormatState, group: GroupId): FormatState {
  return registry.options.filter((o) => o.group === group).reduce((acc, o) => resetOption(registry, acc, o.id), state);
}

/**
 * Đối chiếu giá trị đã lưu (cài đặt, preset) với registry hiện hành: giữ giá trị hợp lệ, điền mặc định
 * cho tuỳ chọn còn thiếu, bỏ giá trị lỗi thời và cho biết lý do. Giá trị ngoài quy định chỉ được giữ
 * nếu trước đó đã được xác nhận là ngoại lệ chủ ý.
 */
export function reconcile(registry: FormatRegistry, saved: Record<string, unknown>, savedExceptions: readonly string[]): { state: FormatState; ignored: string[] } {
  const state = initialState(registry);
  const ignored: string[] = [];
  for (const [id, value] of Object.entries(saved)) {
    const def = registry.options.find((o) => o.id === id);
    if (!def) {
      ignored.push(`Tuỳ chọn "${id}" không còn trong chuẩn ${registry.standardName}, đã bỏ qua.`);
      continue;
    }
    const check = checkValue(def, value);
    if (check.status === "invalid") {
      ignored.push(`${check.reason} Đã dùng giá trị mặc định.`);
    } else if (check.status === "exception" && !savedExceptions.includes(id)) {
      ignored.push(`${check.reason} Giá trị chưa được xác nhận, đã dùng mặc định.`);
    } else {
      state.values[id] = value as OptionValue;
      state.exceptions = withException(state.exceptions, id, check.status === "exception");
    }
  }
  return { state, ignored };
}

function num(values: FormatValues, id: string, fallback: number): number {
  const value = values[id];
  return typeof value === "number" ? value : fallback;
}

/** Giá trị tuỳ chọn → thông số dựng bộ style của chuẩn. */
export function toStyleOptions(profile: StandardProfile, values: FormatValues): StyleOptions {
  const base = defaultStyleOptions(profile);
  const spacing = num(values, "spacing", base.spaceBeforePt);
  const lineValue = num(values, "lineSpacing", base.lineSpacing.kind === "multiple" ? base.lineSpacing.lines : base.lineSpacing.pt);
  const tableSpacing = num(values, "tableSpacing", profile.styles.table.spaceBeforePt);
  return {
    fontSizePt: num(values, "fontSize", base.fontSizePt),
    spaceBeforePt: spacing,
    spaceAfterPt: spacing,
    firstLineIndentMm: num(values, "firstLineIndent", base.firstLineIndentMm),
    lineSpacing: profile.content.lineSpacing.kind === "multiple" ? { kind: "multiple", lines: lineValue } : { kind: "exactly", pt: lineValue },
    alignment: values.alignment === "Left" ? "Left" : "Justified",
    table: { fontSizePt: num(values, "tableFontSize", profile.styles.table.fontSizePt), spaceBeforePt: tableSpacing, spaceAfterPt: tableSpacing },
  };
}

/** Giá trị tuỳ chọn → thiết lập trang (A4 dọc). */
export function toPageSetupSettings(values: FormatValues): PageSetupSettings {
  return {
    paperSize: "A4",
    orientation: "Portrait",
    marginsMm: {
      top: num(values, "marginTop", 20),
      bottom: num(values, "marginBottom", 20),
      left: num(values, "marginLeft", 30),
      right: num(values, "marginRight", 20),
    },
    headerDistanceMm: num(values, "headerDistance", 12.5),
    footerDistanceMm: num(values, "footerDistance", 12.5),
    mirrorMargins: values.duplex === true,
  };
}

export function behaviorOf(values: FormatValues): { localFormatting: "keep" | "clear" } {
  return {
    localFormatting: values.localFormatting === "clear" ? "clear" : "keep",
  };
}
