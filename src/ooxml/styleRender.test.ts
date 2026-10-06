import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { expectMatchesGolden } from "./__golden__/golden";
import { hasExactLineRule, readJustification, readLineRule, styleDefinition } from "./facts";
import { toStyleImportJson } from "./styleJson";
import { buildStyleSpecs } from "./styleSpecs";
import { styleIdFromName, toStyleDefinitionsXml } from "./styleXml";

/** Kiểm tra XML đóng mở thẻ cân đối, đủ để phát hiện lỗi dựng chuỗi. */
function isBalanced(xml: string): boolean {
  const stack: string[] = [];
  const tag = /<(\/?)([A-Za-z][\w:.-]*)([^>]*?)(\/?)>/g;
  let match: RegExpExecArray | null;
  while ((match = tag.exec(xml)) !== null) {
    const [, closing, name, , selfClosing] = match;
    if (selfClosing) continue;
    if (closing) {
      if (stack.pop() !== name) return false;
    } else {
      stack.push(name);
    }
  }
  return stack.length === 0;
}

describe("toStyleImportJson", () => {
  const nd30 = JSON.parse(toStyleImportJson(buildStyleSpecs(ND30)));
  const hd05 = JSON.parse(toStyleImportJson(buildStyleSpecs(HD05)));
  const find = (json: { styles: Array<{ nameLocal: string }> }, name: string) => {
    const found = json.styles.find((s) => s.nameLocal === name);
    if (!found) throw new Error(`Không có ${name}`);
    return found as Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
  };

  it("mỗi style là kiểu đoạn, kế thừa Normal, không phải style có sẵn", () => {
    for (const s of [...nd30.styles, ...hd05.styles]) {
      expect(s).toMatchObject({ type: "Paragraph", baseStyle: "Normal", builtIn: false });
    }
    expect(nd30.styles).toHaveLength(buildStyleSpecs(ND30).length);
  });

  it("ND30_NoiDung: canh đều, thụt 36 pt, giãn 6/6, giãn dòng 1,17 dòng = 14,04 pt, Times New Roman 14", () => {
    const s = find(nd30, "ND30_NoiDung");
    expect(s.paragraphFormat).toMatchObject({ alignment: "Justified", firstLineIndent: 36, spaceBefore: 6, spaceAfter: 6, lineSpacing: 14.04 });
    expect(s.font).toMatchObject({ name: "Times New Roman", size: 14, bold: false, italic: false, color: "#000000", underline: "None" });
  });

  it("HD05_NoiDung: giãn dòng 20 pt", () => {
    expect(find(hd05, "HD05_NoiDung").paragraphFormat.lineSpacing).toBe(20);
  });

  it("giữ với đoạn sau cho tiêu đề và Điều; chữ đậm, nghiêng theo quy tắc", () => {
    expect(find(nd30, "ND30_Dieu").paragraphFormat.keepWithNext).toBe(true);
    expect(find(nd30, "ND30_Dieu").font.bold).toBe(true);
    expect(find(nd30, "ND30_CanCu").font.italic).toBe(true);
    expect(find(nd30, "ND30_NoiDung").paragraphFormat.keepWithNext).toBe(false);
  });

  it("khớp file mẫu đối chiếu", () => {
    expectMatchesGolden("nd30-styles.json", toStyleImportJson(buildStyleSpecs(ND30)));
    expectMatchesGolden("hd05-styles.json", toStyleImportJson(buildStyleSpecs(HD05)));
  });
});

describe("styleIdFromName", () => {
  it("bỏ dấu gạch dưới như Word khi tạo mã style từ tên", () => {
    expect(styleIdFromName("ND30_NoiDung")).toBe("ND30NoiDung");
    expect(styleIdFromName("HD05_TenPhanChuong")).toBe("HD05TenPhanChuong");
  });
});

describe("toStyleDefinitionsXml", () => {
  const nd30 = toStyleDefinitionsXml(buildStyleSpecs(ND30));
  const hd05 = toStyleDefinitionsXml(buildStyleSpecs(HD05));

  it("XML cân đối và có mỗi style một lần, mã theo quy tắc của Word", () => {
    expect(isBalanced(nd30)).toBe(true);
    expect(isBalanced(hd05)).toBe(true);
    expect(nd30.match(/<w:style /g)).toHaveLength(buildStyleSpecs(ND30).length);
    expect(nd30).toContain('w:styleId="ND30NoiDung"');
    expect(nd30).toContain('<w:name w:val="ND30_NoiDung"/>');
  });

  it("ND30_NoiDung: giãn dòng bội số 281 (1,17 x 240), canh đều, thụt 720 twips, cỡ 14", () => {
    const def = styleDefinition(nd30, "ND30NoiDung");
    expect(readLineRule(def)).toBe("auto");
    expect(def).toContain('w:line="281"');
    expect(readJustification(def)).toBe("both");
    expect(def).toContain('<w:ind w:firstLine="720"/>');
    expect(def).toContain('<w:sz w:val="28"/>');
    expect(def).toContain('w:ascii="Times New Roman"');
  });

  it("chuẩn Đảng: giãn dòng Exactly 20 pt = 400 twips", () => {
    expect(hasExactLineRule(styleDefinition(hd05, "HD05NoiDung"), 400)).toBe(true);
    expect(readLineRule(styleDefinition(hd05, "HD05BangDuLieu"))).toBe("auto");
  });

  it("chữ đậm, giữ với đoạn sau và canh giữa thể hiện đúng", () => {
    const dieu = styleDefinition(nd30, "ND30Dieu");
    expect(dieu).toContain("<w:b/>");
    expect(dieu).toContain("<w:keepNext/>");
    expect(readJustification(styleDefinition(nd30, "ND30KinhGui"))).toBe("center");
    expect(styleDefinition(nd30, "ND30NoiDung")).not.toContain("<w:keepNext/>");
  });

  it("khớp file mẫu đối chiếu", () => {
    expectMatchesGolden("nd30-styles.xml", nd30);
    expectMatchesGolden("hd05-styles.xml", hd05);
  });
});
