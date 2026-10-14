import { extractDocDefaults } from "../ooxml/flatOpc";

/** docDefaults của tài liệu đang mở, để đưa vào gói insertOoxml (xem `flatOpc`); undefined nếu đọc không được. */
export async function readDocDefaults(context: Word.RequestContext): Promise<string | undefined> {
  const ooxml = context.document.body.getOoxml();
  await context.sync();
  return extractDocDefaults(ooxml.value);
}
