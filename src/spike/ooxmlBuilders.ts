/**
 * Dựng OOXML thử nghiệm cho spike T3. Các chuỗi này mô phỏng điều T7 đến T11 sẽ làm thật:
 * đường kẻ bằng viền đoạn, bảng đầu văn bản không viền, style, số trang, ngắt phần đánh số lại.
 */

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const FONT = `<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>`;

export function mmToTwips(mm: number): number {
  return Math.round((mm * 1440) / 25.4);
}

function sizeXml(pt: number): string {
  return `<w:sz w:val="${pt * 2}"/><w:szCs w:val="${pt * 2}"/>`;
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
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="${W_NS}"><w:body>${bodyXml}</w:body></w:document></pkg:xmlData></pkg:part>` +
    stylesPart +
    `</pkg:package>`
  );
}

/** Đoạn rỗng có viền dưới, thụt hai bên để đường kẻ ngắn hơn bề rộng dòng chữ và nằm giữa. */
export function lineParagraphBody(options: { leftTwips: number; rightTwips: number }): string {
  return (
    `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>` +
    `<w:spacing w:before="0" w:after="0" w:line="120" w:lineRule="exact"/>` +
    `<w:ind w:left="${options.leftTwips}" w:right="${options.rightTwips}"/><w:jc w:val="center"/></w:pPr></w:p>`
  );
}

function cellParagraph(text: string, options: { bold?: boolean; italic?: boolean; sizePt?: number } = {}): string {
  const { bold = false, italic = false, sizePt = 13 } = options;
  const run = text
    ? `<w:r><w:rPr>${FONT}${bold ? "<w:b/>" : ""}${italic ? "<w:i/>" : ""}${sizeXml(sizePt)}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`
    : "";
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="center"/></w:pPr>${run}</w:p>`;
}

function cell(widthTwips: number, paragraphs: string[]): string {
  return `<w:tc><w:tcPr><w:tcW w:w="${widthTwips}" w:type="dxa"/></w:tcPr>${paragraphs.join("")}</w:tc>`;
}

/** Bảng 3 hàng, 2 cột không viền của khối đầu văn bản (cơ quan | quốc hiệu, số | ngày, V/v | trống). */
export function headerTableBody(options: { totalWidthTwips: number; leftRatio: number; indentTwips: number }): string {
  const { totalWidthTwips, leftRatio, indentTwips } = options;
  const left = Math.round(totalWidthTwips * leftRatio);
  const right = totalWidthTwips - left;
  const none = (side: string) => `<w:${side} w:val="none" w:sz="4" w:space="0" w:color="auto"/>`;
  const borders = ["top", "left", "bottom", "right", "insideH", "insideV"].map(none).join("");
  const rows = [
    [
      [cellParagraph("TÊN CƠ QUAN CHỦ QUẢN"), cellParagraph("TÊN CƠ QUAN BAN HÀNH", { bold: true })],
      [cellParagraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", { bold: true }), cellParagraph("Độc lập - Tự do - Hạnh phúc", { bold: true, sizePt: 14 })],
    ],
    [[cellParagraph("Số: 05/XX-YY")], [cellParagraph("Hà Nội, ngày 05 tháng 01 năm 2026", { italic: true })]],
    [[cellParagraph("V/v thử nghiệm bố cục khối đầu văn bản")], [cellParagraph("")]],
  ];
  const rowXml = rows.map(([l, r]) => `<w:tr>${cell(left, l)}${cell(right, r)}</w:tr>`).join("");
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="${totalWidthTwips}" w:type="dxa"/><w:tblInd w:w="${indentTwips}" w:type="dxa"/>` +
    `<w:tblBorders>${borders}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="${left}"/><w:gridCol w:w="${right}"/></w:tblGrid>${rowXml}</w:tbl>`
  );
}

export interface StyleDefinitionOptions {
  id: string;
  name: string;
  fontPt: number;
  /** Giãn dòng kiểu Exactly, pt. */
  exactLineSpacingPt?: number;
}

/** Định nghĩa một style đoạn (phần tử w:style) để đưa vào phần styles của gói OOXML. */
export function styleDefinitionXml(options: StyleDefinitionOptions): string {
  const { id, name, fontPt, exactLineSpacingPt } = options;
  const line = exactLineSpacingPt === undefined ? "" : ` w:line="${exactLineSpacingPt * 20}" w:lineRule="exact"`;
  return (
    `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:qFormat/>` +
    `<w:pPr><w:spacing w:before="120" w:after="120"${line}/><w:ind w:firstLine="567"/><w:jc w:val="both"/></w:pPr>` +
    `<w:rPr>${FONT}${sizeXml(fontPt)}</w:rPr></w:style>`
  );
}

export interface StyleImportOptions {
  name: string;
  fontPt: number;
  /** Giãn dòng theo pt như API Word.ParagraphFormat.lineSpacing. */
  lineSpacingPt: number;
}

/** Chuỗi JSON cho document.importStylesFromJson, theo dạng ví dụ của Microsoft. */
export function styleImportJson(options: StyleImportOptions): string {
  return JSON.stringify({
    styles: [
      {
        baseStyle: "Normal",
        builtIn: false,
        inUse: true,
        linked: false,
        nameLocal: options.name,
        priority: 1,
        quickStyle: true,
        type: "Paragraph",
        unhideWhenUsed: false,
        visibility: false,
        paragraphFormat: {
          alignment: "Justified",
          firstLineIndent: 28.35,
          keepTogether: false,
          keepWithNext: false,
          leftIndent: 0,
          lineSpacing: options.lineSpacingPt,
          lineUnitAfter: 0,
          lineUnitBefore: 0,
          mirrorIndents: false,
          outlineLevel: "OutlineLevelBodyText",
          rightIndent: 0,
          spaceAfter: 6,
          spaceBefore: 6,
          widowControl: true,
        },
        font: {
          name: "Times New Roman",
          size: options.fontPt,
          bold: false,
          italic: false,
          color: "#000000",
          underline: "None",
          subscript: false,
          superscript: false,
          strikeThrough: false,
          doubleStrikeThrough: false,
          highlightColor: null,
          hidden: false,
        },
      },
    ],
  });
}

/** Đoạn canh giữa chứa trường số trang PAGE. */
export function pageNumberParagraphBody(fontPt: number): string {
  return (
    `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>` +
    `<w:fldSimple w:instr=" PAGE "><w:r><w:rPr>${FONT}${sizeXml(fontPt)}</w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p>`
  );
}

/** Đoạn kết thúc một phần, phần đó bắt đầu trang mới và đánh số trang lại từ 1 (phụ lục). */
export function sectionBreakRestartBody(): string {
  return `<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/><w:pgNumType w:start="1"/></w:sectPr></w:pPr></w:p>`;
}
