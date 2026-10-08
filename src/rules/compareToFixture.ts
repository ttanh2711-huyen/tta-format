import type { DocModel } from "./docModel";

/** Sai số khi so số đo (mm, pt, bội số) sau khi đổi đơn vị và làm tròn. */
const TOLERANCE = 0.06;

export interface CompareOptions {
  /** Đường dẫn (ví dụ "blocks[1].columnsMm") không so sánh. */
  ignore?: readonly RegExp[];
}

type Json = unknown;

function show(value: Json): string {
  return typeof value === "string" ? `"${value}"` : String(value);
}

/** Chữ của mẫu đã ẩn danh có dạng "Nhãn: […]" hoặc "[…]": chỉ so nhãn đầu dòng. */
function compareText(path: string, actual: string, expected: string, out: string[]): void {
  const placeholder = expected.indexOf("[…]");
  if (placeholder < 0) {
    if (actual !== expected) out.push(`${path}: tài liệu ${show(actual)}, mẫu ${show(expected)}`);
    return;
  }
  const label = expected.slice(0, placeholder).trim();
  if (label !== "" && !actual.trim().startsWith(label)) {
    out.push(`${path}: tài liệu ${show(actual)}, mẫu bắt đầu bằng ${show(label)}`);
  }
}

function walk(path: string, actual: Json, expected: Json, ignore: readonly RegExp[], out: string[]): void {
  if (ignore.some((pattern) => pattern.test(path))) return;
  if (Array.isArray(expected)) {
    if (!Array.isArray(actual)) {
      out.push(`${path}: tài liệu ${show(actual)}, mẫu là danh sách`);
      return;
    }
    if (actual.length !== expected.length) {
      out.push(`${path}: tài liệu có ${actual.length} phần tử, mẫu có ${expected.length}`);
      return;
    }
    expected.forEach((item, index) => walk(`${path}[${index}]`, actual[index], item, ignore, out));
    return;
  }
  if (expected !== null && typeof expected === "object") {
    const actualObject = (actual ?? {}) as Record<string, Json>;
    for (const [key, value] of Object.entries(expected)) {
      const child = path === "" ? key : `${path}.${key}`;
      if (ignore.some((pattern) => pattern.test(child))) continue;
      if (!(key in actualObject) || actualObject[key] === undefined) {
        out.push(`${child}: tài liệu không có, mẫu ${typeof value === "object" ? "có" : show(value)}`);
        continue;
      }
      if (key === "text" && typeof value === "string" && typeof actualObject[key] === "string") {
        compareText(child, actualObject[key], value, out);
      } else {
        walk(child, actualObject[key], value, ignore, out);
      }
    }
    return;
  }
  if (typeof expected === "number" && typeof actual === "number") {
    if (Math.abs(actual - expected) > TOLERANCE) out.push(`${path}: tài liệu ${actual}, mẫu ${expected}`);
    return;
  }
  if (actual !== expected) out.push(`${path}: tài liệu ${show(actual)}, mẫu ${show(expected)}`);
}

/**
 * So mô hình tài liệu với fixture (đáp án chuẩn). Chỉ so các thuộc tính mà fixture khai báo, nên tài
 * liệu có thể có thêm thuộc tính khác. Trả về danh sách khác biệt bằng tiếng Việt, rỗng nếu khớp.
 */
export function compareToFixture(actual: DocModel, fixture: DocModel, options: CompareOptions = {}): string[] {
  const out: string[] = [];
  walk("", actual, fixture, options.ignore ?? [], out);
  return out;
}
