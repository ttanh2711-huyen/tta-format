import { flatOpc } from "../ooxml/flatOpc";
import { paragraphStyleIds, readLineRule, styleDefinition, styleIdByName } from "../ooxml/facts";
import { toStyleImportJson } from "../ooxml/styleJson";
import type { StyleSpec } from "../ooxml/styleSpecs";
import { styleIdFromName, toStyleDefinitionsXml } from "../ooxml/styleXml";
import { planExactRepair } from "./exactStyles";
import { findStyleMismatches, type StyleInstallItem, type StyleReadback } from "./styleVerify";

export interface InstallResult {
  items: StyleInstallItem[];
  /** Ghi chú chung, ví dụ thiếu WordApiDesktop 1.1 nên không cập nhật được style đã có. */
  notes: string[];
}

/**
 * Cài style: có WordApi 1.6 (Microsoft 365) thì dùng importStylesFromJson và cập nhật được style đã có; Word 2016, 2019,
 * 2021 (WordApi 1.3) thì thêm style còn thiếu bằng OOXML (style đã có giữ nguyên).
 */
export function isStyleInstallSupported(): boolean {
  return Office.context.requirements.isSetSupported("WordApi", "1.3");
}

const hasFullStyleApi = (): boolean => Office.context.requirements.isSetSupported("WordApi", "1.6");

export const LEGACY_STYLE_NOTE = "Bản Word này chưa có API cài style (cần Microsoft 365): add-in thêm các style còn thiếu bằng OOXML, style đã có trong tài liệu giữ nguyên giá trị cũ.";

interface StylesRead {
  readbacks: Map<string, StyleReadback>;
  /** Mã style đang có đoạn dùng trong thân văn bản. */
  usedIds: Set<string>;
}

async function readStyles(context: Word.RequestContext, specs: readonly StyleSpec[]): Promise<StylesRead> {
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
  return { readbacks: result, usedIds: new Set(paragraphStyleIds(ooxml.value)) };
}

/**
 * Cài style bằng OOXML để có giãn dòng Exactly. Word chỉ nhập style mà nội dung chèn có dùng, nên chèn mỗi style một
 * đoạn mẫu ở đầu thân rồi xoá đoạn mẫu đi; style ở lại trong tài liệu.
 */
async function insertStylesFromOoxml(context: Word.RequestContext, specs: readonly StyleSpec[]): Promise<void> {
  const body = context.document.body;
  const paragraphs = body.paragraphs;
  paragraphs.load("items");
  await context.sync();
  const countBefore = paragraphs.items.length;
  const bodyXml = specs.map((spec) => `<w:p><w:pPr><w:pStyle w:val="${styleIdFromName(spec.name)}"/></w:pPr><w:r><w:t>.</w:t></w:r></w:p>`).join("");
  const inserted = paragraphs.getFirst().getRange("Start").insertOoxml(flatOpc({ bodyXml, stylesXml: toStyleDefinitionsXml(specs) }), "Replace");
  await context.sync();
  inserted.delete();
  await context.sync();
  // Đoạn mẫu không được để lại đoạn rỗng thừa ở đầu tài liệu.
  const after = body.paragraphs;
  after.load("items/text");
  await context.sync();
  for (let extra = after.items.length - countBefore, i = 0; extra > 0 && i < after.items.length && after.items[i].text.trim() === ""; extra -= 1, i += 1) after.items[i].delete();
  await context.sync();
}

/**
 * Cài hoặc cập nhật bộ style trong một lượt, rồi đọc lại từng style để kiểm tra. Idempotent: chạy
 * lại thì cập nhật tại chỗ (Overwrite), không tạo bản sao, đoạn đã dùng style giữ nguyên tham chiếu.
 * Không đụng tới style Normal hay Heading vì chỉ nhập các style do spec khai báo.
 *
 * Giới hạn đã biết (spike T3): importStylesFromJson cho ra quy tắc giãn dòng atLeast, không phải
 * Exactly; kết quả đọc lại sẽ báo rõ điều đó cho các style của chuẩn Đảng.
 */
export async function installStyles(specs: readonly StyleSpec[], options: { bodyWillBeReplaced?: boolean } = {}): Promise<InstallResult> {
  if (!isStyleInstallSupported()) {
    throw new Error("Cài style cần Word 2016 trở lên. Hãy cập nhật Word rồi mở lại.");
  }
  if (!hasFullStyleApi()) {
    return Word.run(async (context) => {
      await insertStylesFromOoxml(context, specs);
      return { items: specs.map((spec) => ({ name: spec.name, action: "created" as const, mismatches: [] })), notes: [LEGACY_STYLE_NOTE] };
    });
  }
  return Word.run(async (context) => {
    const notes: string[] = [];
    const { readbacks: before, usedIds } = await readStyles(context, specs);
    const canOverwrite = Office.context.requirements.isSetSupported("WordApiDesktop", "1.1");
    // Giãn dòng Exactly chỉ đặt được khi cài mới bằng OOXML: xoá style sai quy tắc rồi dựng lại (xem exactStyles).
    const repair = planExactRepair(specs, before, usedIds, options.bodyWillBeReplaced === true);
    if (repair.rebuild.length > 0) {
      if (options.bodyWillBeReplaced) context.document.body.clear();
      for (const spec of repair.rebuild) {
        if (before.get(spec.name)?.exists) context.document.getStyles().getByNameOrNullObject(spec.name).delete();
      }
      await context.sync();
    }
    if (repair.blocked.length > 0) notes.push(`Không đặt được giãn dòng Exactly cho ${repair.blocked.map((spec) => spec.name).join(", ")} vì đang có đoạn dùng; bấm "Định dạng nội dung" để đặt trực tiếp.`);
    const rebuilding = new Set(repair.rebuild.map((spec) => spec.name));
    const importable = specs.filter((spec) => !rebuilding.has(spec.name));
    if (importable.length > 0) {
      const json = toStyleImportJson(importable);
      if (canOverwrite) {
        context.document.importStylesFromJson(json, "Overwrite");
      } else {
        context.document.importStylesFromJson(json);
        notes.push("Thiếu WordApiDesktop 1.1 nên style đã có không được cập nhật, chỉ thêm style mới.");
      }
      await context.sync();
    }
    if (repair.rebuild.length > 0) await insertStylesFromOoxml(context, repair.rebuild);

    const { readbacks: after } = await readStyles(context, specs);
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
  if (!hasFullStyleApi()) return { toCreate: specs.length, toUpdate: 0 };
  return Word.run(async (context) => {
    const { readbacks: current, usedIds } = await readStyles(context, specs);
    const rebuilding = new Set(planExactRepair(specs, current, usedIds, false).rebuild.map((spec) => spec.name));
    let toCreate = 0;
    let toUpdate = 0;
    for (const spec of specs) {
      const readback = current.get(spec.name);
      if (!readback || !readback.exists) {
        toCreate += 1;
      } else if (rebuilding.has(spec.name) || findStyleMismatches(spec, readback).some((m) => !KNOWN_LIMIT.test(m))) {
        toUpdate += 1;
      }
    }
    return { toCreate, toUpdate };
  });
}
