import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { STANDARDS } from "../standards";
import { vbaFiles } from "./generate";
import { stripDiacritics, toAsciiVba } from "./vbaText";

const DIR = join(__dirname, "..", "..", "vba");
const files = vbaFiles();

/** Giải mã như hàm U của VBA. */
const decode = (s: string): string => s.replace(/\\u([0-9A-F]{4})/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)));

describe("toAsciiVba", () => {
  it("đổi chuỗi có dấu thành U(...) và bỏ dấu chú thích", () => {
    expect(toAsciiVba('Msg "Đã xong" \' Chú thích')).toBe('Msg U("\\u0110\\u00E3 xong") \' Chu thich');
  });
  it("giữ nguyên chuỗi ASCII, kể cả dấu nháy kép đôi", () => {
    expect(toAsciiVba('x = "a""b" & "\\u"')).toBe('x = "a""b" & "\\u"');
  });
  it("giải mã lại đúng chữ gốc", () => {
    const out = toAsciiVba('s = "Cộng hòa xã hội chủ nghĩa Việt Nam"');
    expect(decode(out.slice(out.indexOf('U("') + 3, out.lastIndexOf('")')))).toBe("Cộng hòa xã hội chủ nghĩa Việt Nam");
  });
  it("bỏ dấu tiếng Việt", () => {
    expect(stripDiacritics("Định dạng Đảng")).toBe("Dinh dang Dang");
  });
});

describe("bộ macro VBA", () => {
  it.each(Object.keys(files))("%s chỉ có ký tự ASCII, xuống dòng CRLF, dòng không quá 1000 ký tự", (name) => {
    const text = files[name];
    expect(text.replace(/[\t\r\n]/g, "")).toMatch(/^[ -~]*$/);
    expect(text.split("\r\n").every((line) => !line.includes("\n") && line.length <= 1000)).toBe(true);
  });

  it("có đủ macro người dùng chạy", () => {
    for (const name of ["TTA_ChonChuan", "TTA_ThietLapTrang", "TTA_CaiStyle", "TTA_DinhDangNoiDung", "TTA_DanhSoTrang", "TTA_KhoiDau", "TTA_KhoiCuoi", "TTA_KiemTra"]) {
      expect(files["TTA_Format.bas"]).toContain(`Public Sub ${name}()`);
    }
  });

  it("mỗi Sub, Function đều đóng, Select Case, With, If khối cân đối", () => {
    for (const text of Object.values(files)) {
      const lines = text.split("\r\n").map((line) => line.trim());
      const count = (re: RegExp): number => lines.filter((line) => re.test(line)).length;
      expect(count(/^(Public |Private )?(Sub|Function) /)).toBe(count(/^End (Sub|Function)$/));
      expect(count(/^Select Case /)).toBe(count(/^End Select$/));
      expect(count(/^With /)).toBe(count(/^End With$/));
      expect(count(/^If .* Then$/)).toBe(count(/^End If$/));
      expect(count(/^(For|For Each) /)).toBe(count(/^Next/));
    }
  });

  it("cấu hình lấy đúng giá trị từ hồ sơ chuẩn", () => {
    const config = files["TTA_Config.bas"];
    STANDARDS.forEach((profile, i) => {
      const d = profile.pageDefaults;
      expect(config).toContain(`Case ${i + 1}: topMm = ${d.marginsMm.top}: bottomMm = ${d.marginsMm.bottom}: leftMm = ${d.marginsMm.left}: rightMm = ${d.marginsMm.right}: headerMm = ${d.headerDistanceMm}: footerMm = ${d.footerDistanceMm}`);
      expect(config).toContain(`Case ${i + 1}: CfgPrefix = "${profile.styles.prefix}"`);
      expect(decode(config)).toContain(`Case ${i + 1}: CfgTitle = U("${profile.identifyingTitle}")`);
      expect(config).toContain(`Case "${i + 1}|noiDung": CfgRoleStyle = "${profile.styles.prefix}NoiDung"`);
    });
  });

  it("tệp trong thư mục vba/ khớp mã sinh (UPDATE_VBA=1 để ghi lại)", () => {
    for (const [name, text] of Object.entries(files)) {
      const path = join(DIR, name);
      if (process.env.UPDATE_VBA === "1") writeFileSync(path, text, "latin1");
      expect(existsSync(path)).toBe(true);
      expect(readFileSync(path, "latin1")).toBe(text);
    }
  });
});
