/**
 * Bản chụp tài liệu trước khi "Chỉnh toàn bộ" để khôi phục (kết quả spike T3, phép thử p13: `getOoxml` rồi
 * `insertOoxml(..., "Replace")` đưa thân văn bản, định dạng đoạn, bảng và lề về đúng).
 *
 * Phạm vi: thân văn bản, thiết lập trang và hai header (mặc định, trang đầu) của từng phần. Chưa gồm footer và
 * định nghĩa style: các bước hiện có không sửa footer; các style bắt đầu bằng ND30_ hoặc HD05_ do add-in tạo giữ giá trị mới sau khi khôi
 * phục (insertOoxml không cập nhật style đã có). Task nào sửa header/footer phải mở rộng bản chụp.
 * Bản chụp nằm trong bộ nhớ của task pane: đóng task pane là mất.
 */

export interface SectionSnapshot {
  pageWidth: number;
  pageHeight: number;
  topMargin: number;
  bottomMargin: number;
  leftMargin: number;
  rightMargin: number;
  headerDistance: number;
  footerDistance: number;
  mirrorMargins: boolean;
  differentFirst: boolean;
  /** OOXML của header mặc định và header trang đầu (bước Đánh số trang sửa hai header này). */
  primaryHeaderOoxml: string;
  firstHeaderOoxml: string;
}

export interface DocumentSnapshot {
  takenAt: Date;
  bodyOoxml: string;
  bodyText: string;
  sections: SectionSnapshot[];
}

const PAGE_PROPS = "pageWidth,pageHeight,topMargin,bottomMargin,leftMargin,rightMargin,headerDistance,footerDistance,mirrorMargins,differentFirstPageHeaderFooter";

export async function takeSnapshot(): Promise<DocumentSnapshot> {
  return Word.run(async (context) => {
    const body = context.document.body;
    const ooxml = body.getOoxml();
    body.load("text");
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    for (const section of sections.items) section.pageSetup.load(PAGE_PROPS);
    const headers = sections.items.map((section) => ({ primary: section.getHeader("Primary").getOoxml(), first: section.getHeader("FirstPage").getOoxml() }));
    await context.sync();
    return {
      takenAt: new Date(),
      bodyOoxml: ooxml.value,
      bodyText: body.text,
      sections: sections.items.map((section, index) => {
        const ps = section.pageSetup;
        return {
          pageWidth: ps.pageWidth,
          pageHeight: ps.pageHeight,
          topMargin: ps.topMargin,
          bottomMargin: ps.bottomMargin,
          leftMargin: ps.leftMargin,
          rightMargin: ps.rightMargin,
          headerDistance: ps.headerDistance,
          footerDistance: ps.footerDistance,
          mirrorMargins: ps.mirrorMargins,
          differentFirst: ps.differentFirstPageHeaderFooter,
          primaryHeaderOoxml: headers[index].primary.value,
          firstHeaderOoxml: headers[index].first.value,
        };
      }),
    };
  });
}

export interface RestoreResult {
  /** Điểm khác biệt còn lại sau khi khôi phục; rỗng nghĩa là khớp bản chụp. */
  problems: string[];
}

export async function restoreSnapshot(snapshot: DocumentSnapshot): Promise<RestoreResult> {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.insertOoxml(snapshot.bodyOoxml, "Replace");
    await context.sync();

    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    const problems: string[] = [];
    const count = Math.min(sections.items.length, snapshot.sections.length);
    for (let index = 0; index < count; index += 1) {
      const saved = snapshot.sections[index];
      sections.items[index].pageSetup.set({
        topMargin: saved.topMargin,
        bottomMargin: saved.bottomMargin,
        leftMargin: saved.leftMargin,
        rightMargin: saved.rightMargin,
        headerDistance: saved.headerDistance,
        footerDistance: saved.footerDistance,
        mirrorMargins: saved.mirrorMargins,
        differentFirstPageHeaderFooter: saved.differentFirst,
      });
      sections.items[index].getHeader("Primary").insertOoxml(saved.primaryHeaderOoxml, "Replace");
      sections.items[index].getHeader("FirstPage").insertOoxml(saved.firstHeaderOoxml, "Replace");
    }
    await context.sync();
    if (sections.items.length !== snapshot.sections.length) {
      problems.push(`số phần của tài liệu là ${sections.items.length}, bản chụp có ${snapshot.sections.length}; chỉ khôi phục thiết lập trang cho ${count} phần đầu`);
    }

    body.load("text");
    for (const section of sections.items) section.pageSetup.load(PAGE_PROPS);
    await context.sync();
    if (body.text.trim() !== snapshot.bodyText.trim()) {
      problems.push("chữ trong thân văn bản không khớp bản chụp");
    }
    for (let index = 0; index < count; index += 1) {
      const saved = snapshot.sections[index];
      const ps = sections.items[index].pageSetup;
      const off = (["topMargin", "bottomMargin", "leftMargin", "rightMargin"] as const).some((key) => Math.abs(ps[key] - saved[key]) > 0.5);
      if (off) problems.push(`lề của phần ${index + 1} chưa về đúng giá trị cũ`);
    }
    return { problems };
  });
}

export function describeSnapshot(snapshot: DocumentSnapshot): string {
  const time = snapshot.takenAt.toLocaleTimeString("vi-VN", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  return `bản chụp lúc ${time}`;
}
