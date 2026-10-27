import { planHeaderRebuild, type HeaderRebuildPlan } from "../checker/headerRebuild";
import { QUOC_HIEU } from "../ooxml/headerBlock";
import type { StandardProfile } from "../standards/types";
import { insertHeaderBlock, type InsertResult } from "./insertHeader";
import { readCheckDocument } from "./readDocument";

/** Chỉ đọc: nội dung khối đầu cũ đã tách và chuẩn hoá, kèm vùng sẽ xoá. */
export async function previewHeaderRebuild(profile: StandardProfile): Promise<HeaderRebuildPlan | { error: string }> {
  return planHeaderRebuild(await readCheckDocument(), profile);
}

/**
 * Xoá khối đầu cũ (từ đầu văn bản tới hết trích yếu, "Kính gửi"; bảng xoá cả bảng) rồi chèn khối đầu mới `bodyXml`.
 * Vùng cũ được đọc lại ngay trước khi xoá, để không xoá nhầm nếu văn bản đã đổi sau lúc xem trước.
 */
export async function rebuildHeader(profile: StandardProfile, bodyXml: string): Promise<InsertResult> {
  const doc = await readCheckDocument();
  const plan = planHeaderRebuild(doc, profile);
  if ("error" in plan) throw new Error(plan.error);
  const clean = (s: string): string => s.replace(/\s+/g, " ").trim();
  const expected = doc.paragraphs.slice(0, plan.end + 1).map((p) => clean(p.resolved.text));
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/text,items/tableNestingLevel");
    await context.sync();
    const range = paragraphs.items.slice(0, plan.end + 1);
    // An toàn: chữ từng đoạn trong Word phải khớp với nội dung đã đọc (hộp chữ, đối tượng lồng làm lệch thứ tự đoạn).
    if (range.length !== expected.length || range.some((p, i) => clean(p.text) !== expected[i])) {
      throw new Error("Đầu văn bản có hộp chữ hoặc đối tượng lồng nên không xác định chắc vùng khối đầu cũ; chưa xoá gì. Hãy chèn khối đầu mới ở mục Khối đầu rồi xoá khối cũ bằng tay.");
    }
    // Bảng: xoá một lần qua đoạn đầu tiên của mỗi bảng; đoạn ngoài bảng xoá từng đoạn, từ dưới lên.
    const tables = range.filter((p, i) => p.tableNestingLevel > 0 && (i === 0 || range[i - 1].tableNestingLevel === 0)).map((p) => p.parentTable);
    for (const p of [...range].reverse()) if (p.tableNestingLevel === 0) p.delete();
    for (const table of tables) table.delete();
    await context.sync();
  });
  return insertHeaderBlock(bodyXml, profile.headerKind === "dang" ? profile.identifyingTitle : QUOC_HIEU);
}
