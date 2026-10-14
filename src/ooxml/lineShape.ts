/**
 * Đường thẳng nổi (Insert > Shapes > Line) neo trong một đoạn, như file mẫu: nằm dưới dòng chữ của đoạn chứa nó,
 * canh giữa theo cột và đi theo đoạn khi di chuyển. Dùng cho đường kẻ dưới cơ quan ban hành, tiêu ngữ và tiêu
 * đề Đảng. Chỉ phần DrawingML (không có phần dự phòng VML), cần Word 2010 trở lên.
 */

const EMU_PER_MM = 36000;
const EMU_PER_TWIP = 635;
const EMU_PER_PT = 12700;
/** Độ dày nét: 0,75 pt, như đường thẳng mặc định của Word. */
const LINE_WIDTH_EMU = 9525;

/** Hệ số chiều cao một dòng đơn của Times New Roman so với cỡ chữ. */
const SINGLE_LINE_FACTOR = 1.15;
/** Khoảng đệm (pt) giữa chân dòng chữ và đường kẻ, theo file mẫu (đường cách đỉnh đoạn 17,9 pt với chữ 13 pt). */
const GAP_BELOW_TEXT_PT = 3;

export interface LineShapeInput {
  /** Độ dài đường kẻ (mm). */
  lengthMm: number;
  /** Bề rộng của cột chứa đoạn (twips), để canh đường kẻ vào giữa. */
  columnTwips: number;
  /** Cỡ chữ của dòng chữ mà đường kẻ nằm dưới (pt). */
  textSizePt: number;
  /** Số dòng mà chữ của đoạn chiếm (mặc định 1); đường kẻ nằm dưới dòng cuối. */
  textLines?: number;
  /** Mã hình, duy nhất trong tài liệu. */
  id: number;
}

/** Khoảng cách (pt) từ đỉnh đoạn tới đường kẻ. */
export function lineOffsetPt(textSizePt: number, textLines = 1): number {
  return textSizePt * SINGLE_LINE_FACTOR * Math.max(1, textLines) + GAP_BELOW_TEXT_PT;
}

/** Run chứa đường thẳng nổi; chèn vào đầu đoạn mà đường kẻ nằm dưới. */
export function lineShapeRun(input: LineShapeInput): string {
  const length = Math.round(input.lengthMm * EMU_PER_MM);
  const columnEmu = input.columnTwips * EMU_PER_TWIP;
  const left = Math.max(0, Math.round((columnEmu - length) / 2));
  const top = Math.round(lineOffsetPt(input.textSizePt, input.textLines) * EMU_PER_PT);
  const name = `Straight Connector ${input.id}`;
  return (
    `<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>` +
    `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${251658240 + input.id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${left}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${top}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${length}" cy="0"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
    `<wp:docPr id="${input.id}" name="${name}"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${length}" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>` +
    `<a:ln w="${LINE_WIDTH_EMU}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>` +
    `</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>`
  );
}

/** Đổi độ dài đường kẻ cho cột rộng `columnTwips` (mm); không vượt bề rộng cột. */
export function clampLineMm(lengthMm: number, columnTwips: number): number {
  return Math.min(lengthMm, (columnTwips * 25.4) / 1440);
}
