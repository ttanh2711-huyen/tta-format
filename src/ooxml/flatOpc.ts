/** Gói OOXML phẳng (flat OPC) để đưa vào insertOoxml. */

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/** Các không gian tên cần cho đường thẳng nổi (mc, wp, wps); khai báo ở phần tử gốc để insertOoxml đọc được. */
const DRAWING_NS =
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

export function mmToTwips(mm: number): number {
  return Math.round((mm * 1440) / 25.4);
}

/** Gói OOXML phẳng đưa vào insertOoxml; có phần styles khi truyền `stylesXml`. */
export function flatOpc(options: { bodyXml: string; stylesXml?: string }): string {
  const { bodyXml, stylesXml } = options;
  const relsType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const stylesPart = stylesXml
    ? `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"><pkg:xmlData><w:styles xmlns:w="${W_NS}">${stylesXml}</w:styles></pkg:xmlData></pkg:part>`
    : "";
  const documentRels = stylesXml
    ? `<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relsType}/styles" Target="styles.xml"/></Relationships></pkg:xmlData></pkg:part>`
    : "";
  return (
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relsType}/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>` +
    documentRels +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="${W_NS}" ${DRAWING_NS}><w:body>${bodyXml}</w:body></w:document></pkg:xmlData></pkg:part>` +
    stylesPart +
    `</pkg:package>`
  );
}
