import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { buildStyleSpecs, type StyleSpec } from "../ooxml/styleSpecs";
import { findStyleMismatches, readbackFromSpec, summarizeInstall, type StyleReadback } from "./styleVerify";

function specOf(name: string, profile = ND30): StyleSpec {
  const found = buildStyleSpecs(profile).find((s) => s.name === name);
  if (!found) throw new Error(`Không có ${name}`);
  return found;
}

describe("findStyleMismatches", () => {
  const spec = specOf("ND30_NoiDung");
  const same: StyleReadback = readbackFromSpec(spec);

  it("khớp hoàn toàn thì không có chênh lệch", () => {
    expect(findStyleMismatches(spec, same)).toEqual([]);
  });

  it("chấp nhận sai số làm tròn của Word (giãn dòng 14,04 đọc lại 14,05)", () => {
    expect(findStyleMismatches(spec, { ...same, paragraph: { ...same.paragraph, lineSpacing: 14.05 } })).toEqual([]);
  });

  it("không tìm thấy style thì báo duy nhất một lỗi", () => {
    expect(findStyleMismatches(spec, { ...same, exists: false })).toEqual(["không tìm thấy style"]);
  });

  it("báo từng chênh lệch bằng tiếng Việt", () => {
    const mismatches = findStyleMismatches(spec, {
      ...same,
      font: { ...same.font, name: "Arial", size: 12, bold: true },
      paragraph: { ...same.paragraph, alignment: "Left", firstLineIndent: 0, spaceBefore: 0, spaceAfter: 12, lineSpacing: 20, keepWithNext: true },
    });
    const text = mismatches.join(" | ");
    expect(text).toContain("phông chữ");
    expect(text).toContain("cỡ chữ");
    expect(text).toContain("chữ đậm");
    expect(text).toContain("căn lề");
    expect(text).toContain("thụt dòng đầu");
    expect(text).toContain("giãn trước");
    expect(text).toContain("giãn sau");
    expect(text).toContain("giãn dòng");
    expect(text).toContain("giữ với đoạn sau");
  });
});

describe("findStyleMismatches: quy tắc giãn dòng", () => {
  it("chuẩn Chính quyền cần quy tắc auto (bội số)", () => {
    const spec = specOf("ND30_NoiDung");
    const base = readbackFromSpec(spec);
    expect(findStyleMismatches(spec, { ...base, lineRule: "auto" })).toEqual([]);
    expect(findStyleMismatches(spec, { ...base, lineRule: "exact" }).join(" ")).toContain("quy tắc giãn dòng");
  });

  it("chuẩn Đảng cần Exactly: Word ghi atLeast thì báo là chưa đặt được Exactly", () => {
    const spec = specOf("HD05_NoiDung", HD05);
    const base = readbackFromSpec(spec);
    expect(findStyleMismatches(spec, { ...base, lineRule: "exact" })).toEqual([]);
    const text = findStyleMismatches(spec, { ...base, lineRule: "atLeast" }).join(" ");
    expect(text).toContain("Exactly");
    expect(text).toContain("atLeast");
  });

  it("chưa đọc được quy tắc thì không kết luận sai", () => {
    const spec = specOf("HD05_NoiDung", HD05);
    expect(findStyleMismatches(spec, { ...readbackFromSpec(spec), lineRule: undefined })).toEqual([]);
  });
});

describe("summarizeInstall", () => {
  it("báo đủ khi mọi style đúng, nêu số style tạo mới và cập nhật", () => {
    const text = summarizeInstall([
      { name: "ND30_NoiDung", action: "created", mismatches: [] },
      { name: "ND30_KinhGui", action: "updated", mismatches: [] },
      { name: "ND30_Dieu", action: "updated", mismatches: [] },
    ]);
    expect(text.kind).toBe("ok");
    expect(text.text).toBe("Đã cài 3 style: 1 tạo mới, 2 cập nhật.");
  });

  it("cảnh báo và liệt kê style lệch kèm lý do", () => {
    const text = summarizeInstall([
      { name: "HD05_NoiDung", action: "created", mismatches: ["quy tắc giãn dòng: cần Exactly, Word ghi atLeast"] },
      { name: "HD05_Dieu", action: "created", mismatches: [] },
    ]);
    expect(text.kind).toBe("warning");
    expect(text.text).toContain("1/2 style chưa đúng");
    expect(text.text).toContain("HD05_NoiDung");
    expect(text.text).toContain("Exactly");
  });
});
