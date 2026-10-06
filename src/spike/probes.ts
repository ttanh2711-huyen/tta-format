/**
 * Các phép thử khả thi Office.js cho T3. Mỗi phép thử đổi tài liệu rồi **đọc lại** để xác nhận
 * thay đổi đã vào Word, không chỉ tin vào việc API không báo lỗi. Chạy trên tài liệu thử trống.
 */
import { mmToPt, ptToMm } from "../standards/units";
import {
  flatOpc,
  headerTableBody,
  lineParagraphBody,
  mmToTwips,
  pageNumberParagraphBody,
  sectionBreakRestartBody,
  styleDefinitionXml,
  styleImportJson,
} from "./ooxmlBuilders";
import {
  hasExactLineRule,
  hasPageField,
  readJustification,
  readPgNumTypeStart,
  readTableIndentTwips,
  tableBordersAllNone,
} from "./ooxmlFacts";
import type { ProbeGroup } from "./report";

export interface ProbeOutcome {
  status: "pass" | "fail" | "manual" | "skipped";
  detail: string;
  data?: unknown;
}

export interface Requirement {
  set: "WordApi" | "WordApiDesktop";
  version: string;
}

export interface Probe {
  id: string;
  title: string;
  group: ProbeGroup;
  critical: boolean;
  requires?: readonly Requirement[];
  run(): Promise<ProbeOutcome>;
}

/** Gom các kết luận nhỏ của một phép thử: có một điều sai là cả phép thử không đạt. */
class Checks {
  private failures: string[] = [];
  private notes: string[] = [];

  expect(ok: boolean, message: string): void {
    if (!ok) this.failures.push(message);
  }

  near(label: string, actual: number, expected: number, tolerance: number, unit: string): void {
    const ok = Math.abs(actual - expected) <= tolerance;
    this.expect(ok, `${label}: đọc lại ${round(actual)} ${unit}, mong đợi ${round(expected)} ${unit}`);
    if (ok) this.note(`${label} ${round(actual)} ${unit}`);
  }

  note(text: string): void {
    this.notes.push(text);
  }

  outcome(extra = ""): ProbeOutcome {
    const tail = extra ? ` ${extra}` : "";
    if (this.failures.length > 0) {
      return { status: "fail", detail: `${this.failures.join("; ")}.${tail}` };
    }
    return { status: "pass", detail: `${this.notes.join("; ")}.${tail}` };
  }
}

function round(n: number): number {
  return Math.round(n * 100) / 100;
}

function errorText(error: unknown): string {
  if (error instanceof Error) {
    const withCode = error as Error & { code?: string };
    return withCode.code ? `${withCode.code}: ${error.message}` : error.message;
  }
  return String(error);
}

async function optional<T>(action: () => Promise<T>): Promise<{ value?: T; error?: string }> {
  try {
    return { value: await action() };
  } catch (error) {
    return { error: errorText(error) };
  }
}

const A4_MM = { width: 210, height: 297 };

// ---------------------------------------------------------------------------------------------

const pageSetupProbe: Probe = {
  id: "p02",
  title: "Đặt A4 và lề theo mm bằng Word.PageSetup, đọc lại",
  group: "chung",
  critical: true,
  requires: [{ set: "WordApiDesktop", version: "1.3" }],
  async run() {
    return Word.run(async (context) => {
      const checks = new Checks();
      const ps = context.document.pageSetup;
      // Đặt khổ giấy trước, đồng bộ xong mới đặt lề, vì đổi khổ giấy có thể làm Word tính lại lề.
      ps.set({ paperSize: "A4", orientation: "Portrait" });
      await context.sync();
      ps.set({
        topMargin: mmToPt(20),
        bottomMargin: mmToPt(20),
        leftMargin: mmToPt(30),
        rightMargin: mmToPt(20),
      });
      await context.sync();
      ps.load("paperSize,orientation,pageWidth,pageHeight,topMargin,bottomMargin,leftMargin,rightMargin");
      await context.sync();
      checks.expect(ps.paperSize === "A4", `paperSize đọc lại là ${ps.paperSize}`);
      checks.near("rộng", ptToMm(ps.pageWidth), A4_MM.width, 0.5, "mm");
      checks.near("cao", ptToMm(ps.pageHeight), A4_MM.height, 0.5, "mm");
      checks.near("lề trên", ptToMm(ps.topMargin), 20, 0.1, "mm");
      checks.near("lề dưới", ptToMm(ps.bottomMargin), 20, 0.1, "mm");
      checks.near("lề trái", ptToMm(ps.leftMargin), 30, 0.1, "mm");
      checks.near("lề phải", ptToMm(ps.rightMargin), 20, 0.1, "mm");
      return checks.outcome();
    });
  },
};

const sectionPageSetupProbe: Probe = {
  id: "p03",
  title: "Lề riêng từng phần (section): đổi phần 2, phần 1 giữ nguyên",
  group: "chung",
  critical: false,
  requires: [{ set: "WordApiDesktop", version: "1.3" }],
  async run() {
    return Word.run(async (context) => {
      const checks = new Checks();
      const body = context.document.body;
      body.clear();
      body.insertText("Phần 1", "Start");
      body.insertBreak("SectionNext", "End");
      body.insertText("Phần 2", "End");
      await context.sync();
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();
      checks.expect(sections.items.length >= 2, `chỉ có ${sections.items.length} phần sau khi ngắt phần`);
      if (sections.items.length >= 2) {
        sections.items[0].pageSetup.leftMargin = mmToPt(30);
        sections.items[1].pageSetup.leftMargin = mmToPt(15);
        await context.sync();
        sections.items.forEach((s) => s.pageSetup.load("leftMargin"));
        await context.sync();
        checks.near("lề trái phần 1", ptToMm(sections.items[0].pageSetup.leftMargin), 30, 0.1, "mm");
        checks.near("lề trái phần 2", ptToMm(sections.items[1].pageSetup.leftMargin), 15, 0.1, "mm");
        sections.items[1].pageSetup.leftMargin = mmToPt(30);
        await context.sync();
      }
      body.clear();
      await context.sync();
      return checks.outcome();
    });
  },
};

const mirrorMarginsProbe: Probe = {
  id: "p04",
  title: "Chuẩn Đảng: lề đối xứng khi in hai mặt (mirrorMargins)",
  group: "hd05",
  critical: false,
  requires: [{ set: "WordApiDesktop", version: "1.3" }],
  async run() {
    return Word.run(async (context) => {
      const checks = new Checks();
      const ps = context.document.pageSetup;
      ps.set({ leftMargin: mmToPt(30), rightMargin: mmToPt(15), mirrorMargins: true });
      await context.sync();
      ps.load("mirrorMargins,leftMargin,rightMargin");
      await context.sync();
      checks.expect(ps.mirrorMargins === true, "mirrorMargins đọc lại không phải true");
      checks.near("lề trong (leftMargin)", ptToMm(ps.leftMargin), 30, 0.1, "mm");
      checks.near("lề ngoài (rightMargin)", ptToMm(ps.rightMargin), 15, 0.1, "mm");
      ps.mirrorMargins = false;
      await context.sync();
      return checks.outcome("Khi bật, lề trái là lề trong; mặt sau tự đảo thành trái 15, phải 30 (cần xem Print Preview để xác nhận bằng mắt).");
    });
  },
};

const headerDistanceProbe: Probe = {
  id: "p05",
  title: "Khoảng cách header 10 mm và \"trang đầu khác biệt\"",
  group: "hd05",
  critical: false,
  requires: [{ set: "WordApiDesktop", version: "1.3" }],
  async run() {
    return Word.run(async (context) => {
      const checks = new Checks();
      const ps = context.document.pageSetup;
      ps.set({ headerDistance: mmToPt(10), differentFirstPageHeaderFooter: true });
      await context.sync();
      ps.load("headerDistance,differentFirstPageHeaderFooter");
      await context.sync();
      checks.near("khoảng cách header", ptToMm(ps.headerDistance), 10, 0.1, "mm");
      checks.expect(ps.differentFirstPageHeaderFooter === true, "differentFirstPageHeaderFooter đọc lại không phải true");
      ps.set({ headerDistance: mmToPt(12.5), differentFirstPageHeaderFooter: false });
      await context.sync();
      return checks.outcome();
    });
  },
};

const pageNumberProbe: Probe = {
  id: "p06",
  title: "Số trang canh giữa ở header (insertField hoặc OOXML), trang đầu header trống",
  group: "chung",
  critical: true,
  requires: [{ set: "WordApiDesktop", version: "1.3" }],
  async run() {
    return Word.run(async (context) => {
      const checks = new Checks();
      context.document.body.clear();
      const section = context.document.sections.getFirst();
      section.pageSetup.differentFirstPageHeaderFooter = true;
      const header = section.getHeader("Primary");
      header.clear();
      await context.sync();

      const viaField = await optional(async () => {
        const paragraph = header.paragraphs.getFirst();
        paragraph.alignment = "Centered";
        paragraph.getRange("Start").insertField("Start", "Page");
        await context.sync();
        const xml = header.getOoxml();
        await context.sync();
        return xml.value;
      });
      const fieldOk = viaField.value !== undefined && hasPageField(viaField.value) && readJustification(viaField.value) === "center";
      checks.note(`insertField: ${fieldOk ? "đạt" : `không đạt (${viaField.error ?? "không thấy trường PAGE hoặc căn giữa"})`}`);

      header.clear();
      await context.sync();
      const viaOoxml = await optional(async () => {
        header.insertOoxml(flatOpc({ bodyXml: pageNumberParagraphBody(14) }), "Replace");
        await context.sync();
        const xml = header.getOoxml();
        await context.sync();
        return xml.value;
      });
      const ooxmlOk = viaOoxml.value !== undefined && hasPageField(viaOoxml.value) && readJustification(viaOoxml.value) === "center";
      checks.note(`insertOoxml: ${ooxmlOk ? "đạt" : `không đạt (${viaOoxml.error ?? "không thấy trường PAGE hoặc căn giữa"})`}`);

      checks.expect(fieldOk || ooxmlOk, "cả hai cách chèn số trang đều không đạt");

      const first = section.getHeader("FirstPage");
      first.load("text");
      await context.sync();
      checks.expect(first.text.trim() === "", `header trang đầu không trống: "${first.text}"`);
      checks.note("header trang đầu trống");
      return checks.outcome();
    });
  },
};

const lineProbe: Probe = {
  id: "p07",
  title: "Đường kẻ ngắn bằng viền dưới của đoạn rỗng có thụt hai bên",
  group: "chung",
  critical: false,
  async run() {
    return Word.run(async (context) => {
      const checks = new Checks();
      const body = context.document.body;
      body.clear();
      const title = body.insertParagraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Start");
      title.alignment = "Centered";
      title.font.bold = true;
      title.font.name = "Times New Roman";
      title.font.size = 13;
      body.insertOoxml(flatOpc({ bodyXml: lineParagraphBody({ leftTwips: mmToTwips(30), rightTwips: mmToTwips(30) }) }), "End");
      await context.sync();
      const xml = body.getOoxml();
      await context.sync();
      const text = xml.value;
      checks.expect(/<w:pBdr>[\s\S]*?<w:bottom\b[^>]*w:val="single"/.test(text), "không thấy viền dưới trong OOXML của thân văn bản");
      const left = String(mmToTwips(30));
      checks.expect(
        new RegExp(`w:(left|start)="${left}"`).test(text) && new RegExp(`w:(right|end)="${left}"`).test(text),
        `không thấy thụt trái và phải ${left} twips`,
      );
      checks.note("viền dưới và thụt hai bên còn nguyên sau khi vào Word");
      return checks.outcome("Cần xem bằng mắt: đường kẻ phải ngắn hơn dòng chữ phía trên và nằm giữa.");
    });
  },
};

async function insertHeaderTable(
  context: Word.RequestContext,
  options: { totalWidthTwips: number; leftRatio: number; indentTwips: number },
): Promise<{ rows: number; width: number; xml: string }> {
  const body = context.document.body;
  body.clear();
  body.insertOoxml(flatOpc({ bodyXml: headerTableBody(options) }), "Start");
  await context.sync();
  const tables = body.tables;
  tables.load("items/rowCount,items/width");
  await context.sync();
  if (tables.items.length === 0) {
    return { rows: 0, width: 0, xml: "" };
  }
  const xml = tables.items[0].getRange("Whole").getOoxml();
  await context.sync();
  return { rows: tables.items[0].rowCount, width: tables.items[0].width, xml: xml.value };
}

const headerTableProbe: Probe = {
  id: "p08",
  title: "Bảng đầu văn bản 3 hàng x 2 cột không viền bằng insertOoxml, vừa vùng chữ",
  group: "chung",
  critical: true,
  async run() {
    return Word.run(async (context) => {
      const checks = new Checks();
      const fit = await insertHeaderTable(context, { totalWidthTwips: mmToTwips(160), leftRatio: 0.45, indentTwips: 0 });
      checks.expect(fit.rows === 3, `số hàng đọc lại là ${fit.rows}, mong đợi 3`);
      checks.near("bề rộng bảng vừa vùng chữ", fit.width, mmToPt(160), 3, "pt");
      checks.expect(tableBordersAllNone(fit.xml), "OOXML của bảng không còn khai báo 6 cạnh không viền");
      checks.note(`bảng ${fit.rows} hàng, thụt ${readTableIndentTwips(fit.xml) ?? "không ghi"} twips, 6 cạnh không viền`);

      const wide = await insertHeaderTable(context, { totalWidthTwips: 10400, leftRatio: 4700 / 10400, indentTwips: -800 });
      const wideIndent = readTableIndentTwips(wide.xml);
      checks.note(`bảng rộng như file mẫu cũ: rộng ${round(wide.width)} pt, thụt ${wideIndent ?? "không ghi"} twips (mong đợi -800)`);
      context.document.body.clear();
      await context.sync();
      return checks.outcome("Cần xem bằng mắt khi chạy riêng: hai bảng không có đường viền nào.");
    });
  },
};

const apiTableProbe: Probe = {
  id: "p09",
  title: "Bảng tạo bằng API insertTable rồi bỏ viền (phương án thay thế cho OOXML)",
  group: "chung",
  critical: false,
  async run() {
    return Word.run(async (context) => {
      const checks = new Checks();
      const body = context.document.body;
      body.clear();
      const table = body.insertTable(3, 2, "Start", [
        ["TÊN CƠ QUAN", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"],
        ["Số: 05/XX-YY", "Hà Nội, ngày 05 tháng 01 năm 2026"],
        ["V/v thử", ""],
      ]);
      const border = table.getBorder("All");
      border.type = "None";
      await context.sync();
      border.load("type");
      table.load("rowCount");
      await context.sync();
      checks.expect(table.rowCount === 3, `số hàng đọc lại là ${table.rowCount}`);
      checks.expect(border.type === "None", `kiểu viền đọc lại là ${border.type}`);
      checks.note(`bảng ${table.rowCount} hàng, viền "${border.type}"`);
      body.clear();
      await context.sync();
      return checks.outcome();
    });
  },
};

async function readStyle(
  context: Word.RequestContext,
  name: string,
): Promise<{ exists: boolean; font: string; size: number; lineSpacing: number; spaceAfter: number }> {
  const style = context.document.getStyles().getByNameOrNullObject(name);
  style.load("isNullObject");
  style.font.load("name,size");
  style.paragraphFormat.load("lineSpacing,spaceAfter");
  await context.sync();
  if (style.isNullObject) {
    return { exists: false, font: "", size: 0, lineSpacing: 0, spaceAfter: 0 };
  }
  return {
    exists: true,
    font: style.font.name,
    size: style.font.size,
    lineSpacing: style.paragraphFormat.lineSpacing,
    spaceAfter: style.paragraphFormat.spaceAfter,
  };
}

const styleProbe: Probe = {
  id: "p10",
  title: "Cài và cập nhật style: importStylesFromJson (Overwrite) hoặc insertOoxml",
  group: "chung",
  critical: true,
  async run() {
    return Word.run(async (context) => {
      const checks = new Checks();

      // Cách A: importStylesFromJson, rồi nhập lại với Overwrite để cập nhật.
      const nameA = "ND30_ThuA";
      const importA = await optional(async () => {
        context.document.importStylesFromJson(styleImportJson({ name: nameA, fontPt: 14, lineSpacingPt: 20 }));
        await context.sync();
        return readStyle(context, nameA);
      });
      const a1 = importA.value;
      const importedOk = !!a1 && a1.exists && a1.font === "Times New Roman" && a1.size === 14;
      checks.note(
        `importStylesFromJson lần 1: ${importedOk ? `đạt (cỡ ${a1?.size}, giãn dòng ${a1?.lineSpacing} pt)` : `không đạt (${importA.error ?? JSON.stringify(a1)})`}`,
      );

      let overwriteOk = false;
      if (importedOk) {
        const supportsOverwrite = Office.context.requirements.isSetSupported("WordApiDesktop", "1.1");
        if (supportsOverwrite) {
          const again = await optional(async () => {
            context.document.importStylesFromJson(styleImportJson({ name: nameA, fontPt: 13, lineSpacingPt: 22 }), "Overwrite");
            await context.sync();
            return readStyle(context, nameA);
          });
          overwriteOk = !!again.value && again.value.size === 13;
          checks.note(
            `importStylesFromJson Overwrite cập nhật style cũ: ${overwriteOk ? `đạt (cỡ ${again.value?.size}, giãn dòng ${again.value?.lineSpacing} pt)` : `không đạt (${again.error ?? JSON.stringify(again.value)})`}`,
          );
        } else {
          checks.note("importStylesFromJson Overwrite: thiếu WordApiDesktop 1.1");
        }
      }

      // Cách B: insertOoxml kèm phần styles; chèn lần hai với định nghĩa khác để xem có cập nhật không.
      const idB = "NDThuB";
      const nameB = "ND30_ThuB";
      const paragraph = `<w:p><w:pPr><w:pStyle w:val="${idB}"/></w:pPr><w:r><w:t>Thử style</w:t></w:r></w:p>`;
      const insertB = async (fontPt: number, exactPt: number) => {
        context.document.body.insertOoxml(
          flatOpc({ bodyXml: paragraph, stylesXml: styleDefinitionXml({ id: idB, name: nameB, fontPt, exactLineSpacingPt: exactPt }) }),
          "End",
        );
        await context.sync();
        return readStyle(context, nameB);
      };
      const b1 = await optional(() => insertB(14, 20));
      const ooxmlOk = !!b1.value && b1.value.exists && b1.value.size === 14;
      checks.note(`insertOoxml lần 1: ${ooxmlOk ? `đạt (cỡ ${b1.value?.size}, giãn dòng ${b1.value?.lineSpacing} pt)` : `không đạt (${b1.error ?? JSON.stringify(b1.value)})`}`);
      let ooxmlUpdates = false;
      if (ooxmlOk) {
        const b2 = await optional(() => insertB(13, 22));
        ooxmlUpdates = !!b2.value && b2.value.size === 13;
        checks.note(`insertOoxml lần 2 cập nhật style đã có: ${ooxmlUpdates ? "có" : "không (đúng như tài liệu: style đang dùng giữ định nghĩa cũ)"}`);
      }

      checks.expect(importedOk || ooxmlOk, "không cài được style bằng cách nào");
      const body = context.document.body;
      body.clear();
      await context.sync();
      return {
        ...checks.outcome(),
        data: { importedOk, overwriteOk, ooxmlOk, ooxmlUpdates },
      };
    });
  },
};

const exactSpacingProbe: Probe = {
  id: "p11",
  title: "Chuẩn Đảng: giãn dòng Exactly 18-22 pt (style OOXML và API lineSpacing)",
  group: "hd05",
  critical: false,
  async run() {
    return Word.run(async (context) => {
      const checks = new Checks();
      const body = context.document.body;
      body.clear();
      const idC = "NDThuC";
      body.insertOoxml(
        flatOpc({
          bodyXml: `<w:p><w:pPr><w:pStyle w:val="${idC}"/></w:pPr><w:r><w:t>Giãn dòng Exactly 20 pt</w:t></w:r></w:p>`,
          stylesXml: styleDefinitionXml({ id: idC, name: "ND30_ThuC", fontPt: 14, exactLineSpacingPt: 20 }),
        }),
        "Start",
      );
      await context.sync();
      const first = body.paragraphs.getFirst();
      first.load("lineSpacing");
      const styledXml = first.getOoxml();
      await context.sync();
      const exactInStyle = hasExactLineRule(styledXml.value, 400);
      checks.expect(exactInStyle, "OOXML của đoạn không còn giãn dòng Exactly 400 twips (20 pt) trong style");
      checks.note(`style OOXML: Exactly ${exactInStyle ? "còn nguyên" : "mất"}, lineSpacing đọc ra ${first.lineSpacing} pt`);

      const plain = body.insertParagraph("Giãn dòng đặt bằng API lineSpacing = 20", "End");
      plain.lineSpacing = 20;
      await context.sync();
      const plainXml = plain.getOoxml();
      await context.sync();
      const apiExact = hasExactLineRule(plainXml.value, 400);
      const rule = /<w:spacing\b[^>]*w:lineRule="(\w+)"/.exec(plainXml.value)?.[1] ?? "không ghi";
      checks.note(`API lineSpacing=20 cho ra quy tắc "${rule}" (${apiExact ? "là Exactly" : "KHÔNG phải Exactly"})`);
      body.clear();
      await context.sync();
      return { ...checks.outcome(), data: { exactInStyle, apiExact, apiRule: rule } };
    });
  },
};

const appendixNumberingProbe: Probe = {
  id: "p12",
  title: "Phụ lục đánh số trang riêng: ngắt phần có pgNumType start=1 qua insertOoxml",
  group: "chung",
  critical: false,
  async run() {
    return Word.run(async (context) => {
      const checks = new Checks();
      const body = context.document.body;
      body.clear();
      body.insertText("Văn bản chính", "Start");
      body.insertBreak("SectionNext", "End");
      body.insertText("Phụ lục I", "End");
      body.insertOoxml(flatOpc({ bodyXml: sectionBreakRestartBody() }), "End");
      body.insertText("Sau phụ lục", "End");
      await context.sync();
      const sections = context.document.sections;
      sections.load("items");
      const xml = body.getOoxml();
      await context.sync();
      checks.expect(sections.items.length >= 3, `chỉ có ${sections.items.length} phần, mong đợi từ 3`);
      const start = readPgNumTypeStart(xml.value);
      checks.expect(start === 1, `pgNumType start đọc lại là ${start ?? "không có"}, mong đợi 1`);
      checks.note(`${sections.items.length} phần, pgNumType start=${start ?? "không có"}`);
      body.clear();
      await context.sync();
      return checks.outcome("Cần xem bằng mắt: trang đầu của phụ lục phải hiện số 1 ở phần có đánh số.");
    });
  },
};

const snapshotProbe: Probe = {
  id: "p13",
  title: "\"Chỉnh toàn bộ\": chụp bản trước khi chỉnh và khôi phục (thân, bảng, header, lề)",
  group: "chung",
  critical: false,
  requires: [{ set: "WordApiDesktop", version: "1.3" }],
  async run() {
    return Word.run(async (context) => {
      const checks = new Checks();
      const body = context.document.body;
      body.clear();
      const section = context.document.sections.getFirst();
      const header = section.getHeader("Primary");
      header.clear();
      const title = body.insertParagraph("Tiêu đề thử", "Start");
      title.font.set({ name: "Times New Roman", size: 14, bold: true });
      title.alignment = "Centered";
      const content = body.insertParagraph("Nội dung thử nghiệm", "End");
      content.font.set({ name: "Times New Roman", size: 13, italic: true });
      body.insertTable(2, 2, "End", [["a", "b"], ["c", "d"]]);
      header.insertText("Đầu trang thử", "Start");
      section.pageSetup.leftMargin = mmToPt(30);
      await context.sync();

      // Chụp
      const bodySnap = body.getOoxml();
      const headerSnap = header.getOoxml();
      body.paragraphs.load("items/text");
      body.load("text");
      section.pageSetup.load("leftMargin");
      await context.sync();
      const originalText = body.text;
      const originalParagraphs = body.paragraphs.items.length;
      const originalLeft = section.pageSetup.leftMargin;

      // Làm hỏng
      body.paragraphs.load("items");
      await context.sync();
      body.paragraphs.items.forEach((p) => p.font.set({ name: "Arial", size: 10, bold: false, italic: false }));
      header.clear();
      section.pageSetup.leftMargin = mmToPt(15);
      await context.sync();

      // Khôi phục
      body.insertOoxml(bodySnap.value, "Replace");
      header.insertOoxml(headerSnap.value, "Replace");
      section.pageSetup.leftMargin = originalLeft;
      await context.sync();

      // Kiểm tra
      body.load("text");
      body.tables.load("items");
      body.paragraphs.load("items/text,items/font/name,items/font/size,items/font/bold");
      header.load("text");
      section.pageSetup.load("leftMargin");
      await context.sync();
      const firstPara = body.paragraphs.items[0];
      checks.expect(body.text.trim() === originalText.trim(), "văn bản thân khác bản gốc sau khi khôi phục");
      checks.expect(
        !!firstPara && firstPara.font.name === "Times New Roman" && firstPara.font.size === 14 && firstPara.font.bold === true,
        `định dạng đoạn đầu sau khôi phục: ${firstPara?.font.name} ${firstPara?.font.size} đậm=${firstPara?.font.bold}`,
      );
      checks.expect(body.tables.items.length === 1, `số bảng sau khôi phục là ${body.tables.items.length}`);
      checks.expect(header.text.includes("Đầu trang thử"), `header sau khôi phục là "${header.text}"`);
      checks.near("lề trái sau khôi phục", ptToMm(section.pageSetup.leftMargin), ptToMm(originalLeft), 0.1, "mm");
      checks.note(`số đoạn trước ${originalParagraphs}, sau ${body.paragraphs.items.length} (chênh lệch do đoạn rỗng cuối là bình thường)`);
      body.clear();
      header.clear();
      await context.sync();
      return checks.outcome();
    });
  },
};

const undoOneRunProbe: Probe = {
  id: "p14",
  title: "Hoàn tác: 3 thay đổi trong MỘT Word.run, bấm Ctrl+Z bao nhiêu lần?",
  group: "chung",
  critical: false,
  async run() {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.clear();
      ["Đoạn A", "Đoạn B", "Đoạn C"].forEach((t) => body.insertParagraph(t, "End"));
      await context.sync();
    });
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();
      paragraphs.items.slice(0, 3).forEach((p) => {
        p.font.color = "#FF0000";
        p.font.bold = true;
      });
      await context.sync();
    });
    return {
      status: "manual",
      detail:
        "Đã tô đỏ và in đậm 3 đoạn trong MỘT lần Word.run. Bấm Ctrl+Z trong Word: ghi lại cần bấm mấy lần để cả 3 đoạn trở lại bình thường (mong đợi 1).",
    };
  },
};

const undoThreeRunsProbe: Probe = {
  id: "p15",
  title: "Hoàn tác: 3 thay đổi trong BA lần Word.run, bấm Ctrl+Z bao nhiêu lần?",
  group: "chung",
  critical: false,
  async run() {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.clear();
      ["Đoạn A", "Đoạn B", "Đoạn C"].forEach((t) => body.insertParagraph(t, "End"));
      await context.sync();
    });
    for (let i = 0; i < 3; i += 1) {
      await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items");
        await context.sync();
        paragraphs.items[i].font.color = "#FF0000";
        paragraphs.items[i].font.bold = true;
        await context.sync();
      });
    }
    return {
      status: "manual",
      detail: "Đã tô đỏ và in đậm 3 đoạn trong BA lần Word.run riêng. Bấm Ctrl+Z: ghi lại cần bấm mấy lần (để so với phép thử p14).",
    };
  },
};

const performanceProbe: Probe = {
  id: "p16",
  title: "Hiệu năng: 600 đoạn, đặt phông, cỡ, căn đều, giãn dòng trong một lượt",
  group: "chung",
  critical: false,
  async run() {
    return Word.run(async (context) => {
      const body = context.document.body;
      body.clear();
      await context.sync();
      const count = 600;
      const t0 = performance.now();
      for (let i = 1; i <= count; i += 1) {
        body.insertParagraph(`Đoạn ${i}: nội dung thử nghiệm để đo thời gian định dạng hàng loạt trong Word.`, "End");
      }
      await context.sync();
      const t1 = performance.now();
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();
      const t2 = performance.now();
      paragraphs.items.forEach((p) => {
        p.font.name = "Times New Roman";
        p.font.size = 14;
        p.alignment = "Justified";
        p.lineSpacing = 20;
        p.spaceAfter = 6;
        p.firstLineIndent = 28.35;
      });
      await context.sync();
      const t3 = performance.now();
      body.clear();
      await context.sync();
      const insertMs = Math.round(t1 - t0);
      const loadMs = Math.round(t2 - t1);
      const formatMs = Math.round(t3 - t2);
      return {
        status: "pass",
        detail: `${count} đoạn: chèn ${insertMs} ms, tải danh sách ${loadMs} ms, định dạng ${formatMs} ms (khoảng ${Math.round((formatMs / count) * 3000 / 1000)} giây cho 3000 đoạn nếu tuyến tính).`,
        data: { count, insertMs, loadMs, formatMs },
      } as ProbeOutcome;
    });
  },
};

const tabStopProbe: Probe = {
  id: "p17",
  title: "Tab stop (WordApiDesktop 1.4)",
  group: "chung",
  critical: false,
  async run() {
    return {
      status: "skipped",
      detail: "Bỏ qua có chủ đích: các khối đầu và cuối văn bản dựng bằng bảng, không dùng tab stop. Không ảnh hưởng quyết định công nghệ.",
    };
  },
};

export const PROBES: readonly Probe[] = [
  pageSetupProbe,
  sectionPageSetupProbe,
  mirrorMarginsProbe,
  headerDistanceProbe,
  pageNumberProbe,
  lineProbe,
  headerTableProbe,
  apiTableProbe,
  styleProbe,
  exactSpacingProbe,
  appendixNumberingProbe,
  snapshotProbe,
  undoOneRunProbe,
  undoThreeRunsProbe,
  performanceProbe,
  tabStopProbe,
];
