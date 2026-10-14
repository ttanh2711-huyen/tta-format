/** Gói OOXML phẳng (flat OPC) để đưa vào insertOoxml. */

/** Phần `w:docDefaults` trong OOXML (của `body.getOoxml()`), hoặc undefined nếu không có. */
export function extractDocDefaults(ooxml: string): string | undefined {
  const found = /<w:docDefaults>[\s\S]*?<\/w:docDefaults>/.exec(ooxml)?.[0];
  // Bỏ phần tử của không gian tên mở rộng (w14:...), gói của ta chỉ khai báo không gian tên w.
  return found?.replace(/<w14:[^>]*\/>/g, "");
}

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/** Các không gian tên cần cho đường thẳng nổi (mc, wp, wps); khai báo ở phần tử gốc để insertOoxml đọc được. */
const DRAWING_NS =
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

export function mmToTwips(mm: number): number {
  return Math.round((mm * 1440) / 25.4);
}

/** Gói OOXML phẳng đưa vào insertOoxml; có phần styles khi truyền `stylesXml`, phần numbering khi truyền `numberingXml` (cả thẻ `w:numbering`). */
export function flatOpc(options: { bodyXml: string; stylesXml?: string; numberingXml?: string; docDefaultsXml?: string }): string {
  const { bodyXml, numberingXml } = options;
  // Word bỏ các thuộc tính đoạn trực tiếp bằng giá trị mặc định của gói (thụt dòng 0, giãn trước 0) rồi lấy mặc định của
  // tài liệu đích (có thể là thụt dòng đầu 1,27 cm, giãn 6 pt). Cho gói mang đúng docDefaults của tài liệu đích thì các giá trị 0 được giữ.
  const stylesXml = options.stylesXml || options.docDefaultsXml ? `${options.docDefaultsXml ?? ""}${options.stylesXml ?? ""}` : undefined;
  const relsType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const part = (name: string, type: string, xml: string) => `<pkg:part pkg:name="/word/${name}.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`;
  const stylesPart = stylesXml ? part("styles", "styles", `<w:styles xmlns:w="${W_NS}">${stylesXml}</w:styles>`) : "";
  const numberingPart = numberingXml ? part("numbering", "numbering", numberingXml) : "";
  const rels = [stylesXml ? `<Relationship Id="rId1" Type="${relsType}/styles" Target="styles.xml"/>` : "", numberingXml ? `<Relationship Id="rId2" Type="${relsType}/numbering" Target="numbering.xml"/>` : ""].join("");
  const documentRels = rels
    ? `<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships></pkg:xmlData></pkg:part>`
    : "";
  return (
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relsType}/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>` +
    documentRels +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="${W_NS}" ${DRAWING_NS}><w:body>${bodyXml}</w:body></w:document></pkg:xmlData></pkg:part>` +
    stylesPart +
    numberingPart +
    `</pkg:package>`
  );
}
