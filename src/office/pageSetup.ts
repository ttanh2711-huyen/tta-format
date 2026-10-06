import { mmToPt } from "../standards/units";
import { findPageSetupMismatches, type PageSetupSettings } from "./pageSetupPlan";

export interface ApplyOptions {
  /** Đặt phông mặc định (style Normal) thành tên này; bỏ qua nếu không truyền. */
  defaultFontName?: string;
}

export interface ApplyResult {
  sections: number;
  /** Chênh lệch giữa giá trị đã đặt và giá trị Word đọc lại; rỗng nếu khớp. */
  mismatches: string[];
  /** Ghi chú về bước đặt phông mặc định (bước phụ, lỗi ở bước này không làm hỏng thiết lập trang). */
  fontNote?: string;
}

/** Thiết lập trang cần WordApiDesktop 1.3 (Microsoft 365 bản 2507 trở lên). */
export function isPageSetupSupported(): boolean {
  return Office.context.requirements.isSetSupported("WordApiDesktop", "1.3");
}

async function setDefaultFont(context: Word.RequestContext, fontName: string): Promise<string> {
  try {
    const normal = context.document.getStyles().getByNameOrNullObject("Normal");
    normal.load("isNullObject");
    await context.sync();
    if (normal.isNullObject) {
      return "Chưa đặt phông mặc định: không tìm thấy style Normal.";
    }
    normal.font.name = fontName;
    await context.sync();
    normal.font.load("name");
    await context.sync();
    return normal.font.name === fontName
      ? `Phông mặc định: ${fontName}.`
      : `Chưa đặt được phông mặc định: đọc lại là ${normal.font.name}.`;
  } catch (error) {
    return `Chưa đặt phông mặc định: ${error instanceof Error ? error.message : String(error)}`;
  }
}

async function readMismatches(context: Word.RequestContext, items: Word.Section[], settings: PageSetupSettings): Promise<string[]> {
  for (const section of items) {
    section.pageSetup.load(
      "pageWidth,pageHeight,topMargin,bottomMargin,leftMargin,rightMargin,headerDistance,footerDistance,mirrorMargins",
    );
  }
  await context.sync();

  const mismatches: string[] = [];
  items.forEach((section, index) => {
    const ps = section.pageSetup;
    const found = findPageSetupMismatches(settings, {
      widthPt: ps.pageWidth,
      heightPt: ps.pageHeight,
      topPt: ps.topMargin,
      bottomPt: ps.bottomMargin,
      leftPt: ps.leftMargin,
      rightPt: ps.rightMargin,
      headerDistancePt: ps.headerDistance,
      footerDistancePt: ps.footerDistance,
      mirrorMargins: ps.mirrorMargins,
    });
    const prefix = items.length > 1 ? `Phần ${index + 1}: ` : "";
    mismatches.push(...found.map((m) => `${prefix}${m}`));
  });
  return mismatches;
}

/** Chỉ đọc: trang hiện tại có khớp thiết lập không (mỗi phần một lần so). Rỗng nghĩa là đã đúng. */
export async function checkPageSetup(settings: PageSetupSettings): Promise<{ sections: number; mismatches: string[] }> {
  return Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    return { sections: sections.items.length, mismatches: await readMismatches(context, sections.items, settings) };
  });
}

/**
 * Áp dụng kế hoạch lên **mọi phần** của tài liệu rồi đọc lại để xác nhận. Khổ giấy đặt và đồng bộ
 * trước, lề đặt sau, vì đổi khổ giấy có thể làm Word tính lại lề (kết quả spike T3, phép thử p02).
 */
export async function applyPageSetup(settings: PageSetupSettings, options: ApplyOptions = {}): Promise<ApplyResult> {
  return Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    const items = sections.items;

    for (const section of items) {
      section.pageSetup.set({ paperSize: settings.paperSize, orientation: settings.orientation });
    }
    await context.sync();

    for (const section of items) {
      section.pageSetup.set({
        topMargin: mmToPt(settings.marginsMm.top),
        bottomMargin: mmToPt(settings.marginsMm.bottom),
        leftMargin: mmToPt(settings.marginsMm.left),
        rightMargin: mmToPt(settings.marginsMm.right),
        headerDistance: mmToPt(settings.headerDistanceMm),
        footerDistance: mmToPt(settings.footerDistanceMm),
        mirrorMargins: settings.mirrorMargins,
      });
    }
    await context.sync();

    const mismatches = await readMismatches(context, items, settings);

    const fontNote = options.defaultFontName ? await setDefaultFont(context, options.defaultFontName) : undefined;
    return { sections: items.length, mismatches, fontNote };
  });
}
