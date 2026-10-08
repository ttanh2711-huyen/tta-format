const AdmZip = require("adm-zip");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { buildFixture } = require("./extract-sample");

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const body =
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="26"/></w:rPr><w:t>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</w:t></w:r></w:p>' +
  "<w:p><w:r><w:t>Nguyễn Văn A 0909123456 a@example.com</w:t></w:r></w:p>" +
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1701" w:header="709" w:footer="300"/></w:sectPr>';

function makeDocx() {
  const zip = new AdmZip();
  zip.addFile("word/document.xml", Buffer.from(`<w:document ${W}><w:body>${body}</w:body></w:document>`, "utf8"));
  zip.addFile("word/header1.xml", Buffer.from(`<w:hdr ${W}><w:p><w:r><w:instrText> PAGE </w:instrText></w:r></w:p></w:hdr>`, "utf8"));
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "nd30-")), "mau.docx");
  zip.writeZip(file);
  return file;
}

describe("extract-sample", () => {
  it("xuất fixture giữ định dạng và chữ thể thức, bỏ dữ liệu cá nhân", () => {
    const fixture = buildFixture(makeDocx());
    const json = JSON.stringify(fixture);
    expect(fixture.page.marginsMm).toEqual({ top: 20, bottom: 20, left: 30, right: 20 });
    expect(fixture.page.pageNumberInHeader).toBe(true);
    expect(fixture.blocks[0]).toMatchObject({ text: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", align: "center", bold: true, sizePt: 13 });
    expect(json).not.toMatch(/Nguyễn|0909|example\.com/);
  });

  it("hai lần chạy cho cùng kết quả", () => {
    const file = makeDocx();
    expect(JSON.stringify(buildFixture(file))).toBe(JSON.stringify(buildFixture(file)));
  });
});
