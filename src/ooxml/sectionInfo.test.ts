import { pagesFromOoxml } from "./sectionInfo";

const sect = (w: number, h: number, m: { top: number; right: number; bottom: number; left: number; header: number; footer: number }) =>
  `<w:sectPr><w:pgSz w:w="${w}" w:h="${h}"/><w:pgMar w:top="${m.top}" w:right="${m.right}" w:bottom="${m.bottom}" w:left="${m.left}" w:header="${m.header}" w:footer="${m.footer}" w:gutter="0"/></w:sectPr>`;

describe("thiết lập trang đọc từ OOXML (Word 2016, 2019, 2021 không có API thiết lập trang)", () => {
  it("đọc khổ A4 và lề theo mm, đúng thứ tự các phần", () => {
    const xml = `<pkg:package><w:document><w:body><w:p><w:pPr>${sect(11906, 16838, { top: 1134, right: 1134, bottom: 1134, left: 1701, header: 709, footer: 300 })}</w:pPr></w:p><w:p/>${sect(16838, 11906, { top: 1134, right: 1134, bottom: 1134, left: 1134, header: 709, footer: 709 })}</w:body></w:document></pkg:package>`;
    const pages = pagesFromOoxml(xml);
    expect(pages).toHaveLength(2);
    expect(pages[0]).toMatchObject({ heightMm: 297, marginsMm: { top: 20, bottom: 20, left: 30, right: 20 }, headerDistanceMm: 12.51, mirrorMargins: false });
    expect(pages[0].widthMm).toBeCloseTo(210, 1);
    expect(pages[1].widthMm).toBeCloseTo(297, 0);
  });

  it("không có w:sectPr thì không có phần nào", () => {
    expect(pagesFromOoxml("<w:document><w:body><w:p/></w:body></w:document>")).toEqual([]);
  });
});
