import { readLineRule, styleDefinition, styleIdByName } from "../ooxml/facts";
import { toStyleImportJson } from "../ooxml/styleJson";
import type { StyleSpec } from "../ooxml/styleSpecs";
import { findStyleMismatches, type StyleInstallItem, type StyleReadback } from "./styleVerify";

export interface InstallResult {
  items: StyleInstallItem[];
  /** Ghi chú chung, ví dụ thiếu WordApiDesktop 1.1 nên không cập nhật được style đã có. */
  notes: string[];
}

/** Cài style bằng importStylesFromJson cần WordApi 1.6 (Microsoft 365 bản 2308 trở lên). */
export function isStyleInstallSupported(): boolean {
  return Office.context.requirements.isSetSupported("WordApi", "1.6");
}

async function readStyles(context: Word.RequestContext, specs: readonly StyleSpec[]): Promise<Map<string, StyleReadback>> {
  const styles = specs.map((spec) => {
    const style = context.document.getStyles().getByNameOrNullObject(spec.name);
    style.load("isNullObject");
    style.font.load("name,size,bold,italic");
    style.paragraphFormat.load("alignment,firstLineIndent,spaceBefore,spaceAfter,lineSpacing,keepWithNext");
    return { spec, style };
  });
  const ooxml = context.document.body.getOoxml();
  await context.sync();

  const result = new Map<string, StyleReadback>();
  for (const { spec, style } of styles) {
    if (style.isNullObject) {
      result.set(spec.name, {
        exists: false,
        font: { name: "", size: 0, bold: false, italic: false },
        paragraph: { alignment: "", firstLineIndent: 0, spaceBefore: 0, spaceAfter: 0, lineSpacing: 0, keepWithNext: false },
      });
      continue;
    }
    const id = styleIdByName(ooxml.value, spec.name);
    const definition = id ? styleDefinition(ooxml.value, id) : "";
    result.set(spec.name, {
      exists: true,
      font: { name: style.font.name, size: style.font.size, bold: style.font.bold, italic: style.font.italic },
      paragraph: {
        alignment: style.paragraphFormat.alignment,
        firstLineIndent: style.paragraphFormat.firstLineIndent,
        spaceBefore: style.paragraphFormat.spaceBefore,
        spaceAfter: style.paragraphFormat.spaceAfter,
        lineSpacing: style.paragraphFormat.lineSpacing,
        keepWithNext: style.paragraphFormat.keepWithNext,
      },
      lineRule: definition ? readLineRule(definition) : undefined,
    });
  }
  return result;
}

/**
 * Cài hoặc cập nhật bộ style trong một lượt, rồi đọc lại từng style để kiểm tra. Idempotent: chạy
 * lại thì cập nhật tại chỗ (Overwrite), không tạo bản sao, đoạn đã dùng style giữ nguyên tham chiếu.
 * Không đụng tới style Normal hay Heading vì chỉ nhập các style do spec khai báo.
 *
 * Giới hạn đã biết (spike T3): importStylesFromJson cho ra quy tắc giãn dòng atLeast, không phải
 * Exactly; kết quả đọc lại sẽ báo rõ điều đó cho các style của chuẩn Đảng.
 */
export async function installStyles(specs: readonly StyleSpec[]): Promise<InstallResult> {
  if (!isStyleInstallSupported()) {
    throw new Error("Cài style cần Microsoft 365 bản 2308 trở lên (WordApi 1.6). Hãy cập nhật Word rồi mở lại.");
  }
  return Word.run(async (context) => {
    const notes: string[] = [];
    const before = await readStyles(context, specs);
    const canOverwrite = Office.context.requirements.isSetSupported("WordApiDesktop", "1.1");
    const json = toStyleImportJson(specs);
    if (canOverwrite) {
      context.document.importStylesFromJson(json, "Overwrite");
    } else {
      context.document.importStylesFromJson(json);
      notes.push("Thiếu WordApiDesktop 1.1 nên style đã có không được cập nhật, chỉ thêm style mới.");
    }
    await context.sync();

    const after = await readStyles(context, specs);
    const items: StyleInstallItem[] = specs.map((spec) => {
      const existedBefore = before.get(spec.name)?.exists === true;
      const readback = after.get(spec.name);
      return {
        name: spec.name,
        action: existedBefore ? "updated" : "created",
        mismatches: readback ? findStyleMismatches(spec, readback) : ["không đọc lại được"],
      };
    });
    return { items, notes };
  });
}

/** Cách một lần cài sẽ thay đổi tài liệu, tính từ trạng thái hiện tại; chỉ đọc. */
export interface StyleInstallPlan {
  toCreate: number;
  toUpdate: number;
}

/** Quy tắc giãn dòng Exactly không đạt qua importStylesFromJson (spike T3), không tính là "cần cập nhật". */
const KNOWN_LIMIT = /^quy tắc giãn dòng/;

export async function planStyleInstall(specs: readonly StyleSpec[]): Promise<StyleInstallPlan> {
  return Word.run(async (context) => {
    const current = await readStyles(context, specs);
    let toCreate = 0;
    let toUpdate = 0;
    for (const spec of specs) {
      const readback = current.get(spec.name);
      if (!readback || !readback.exists) {
        toCreate += 1;
      } else if (findStyleMismatches(spec, readback).some((m) => !KNOWN_LIMIT.test(m))) {
        toUpdate += 1;
      }
    }
    return { toCreate, toUpdate };
  });
}
