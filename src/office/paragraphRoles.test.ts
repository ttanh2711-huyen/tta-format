import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { assignRoles, detectRole, roleStyle } from "./paragraphRoles";
import { planFormatBody, type ParagraphInfo } from "./paragraphClassifier";

const info = (text: string): ParagraphInfo => ({ text, styleName: "Normal", styleBuiltIn: "Normal", tableNestingLevel: 0, isListItem: false });

describe("nhận diện vai trò đoạn theo ký hiệu đầu đoạn (issue #3)", () => {
  it.each([
    ["Chương I", "chuong"],
    ["CHƯƠNG II", "chuong"],
    ["Chương 3. QUY ĐỊNH CHUNG", "chuong"],
    ["Phần thứ nhất", "chuong"],
    ["PHẦN I", "chuong"],
    ["Mục 1", "muc"],
    ["Mục II. TỔ CHỨC THỰC HIỆN", "muc"],
    ["Điều 5. Phạm vi điều chỉnh", "dieu"],
    ["I. Kết quả thực hiện", "deMuc"],
    ["II. NHIỆM VỤ TRỌNG TÂM", "deMuc"],
    ["1. Tổ chức thực hiện kế hoạch.", "khoan"],
    ["12. Giao Sở Tài chính chủ trì.", "khoan"],
    ["1.1. Về công tác cán bộ", "diem"],
    ["a) Giao Sở Tài chính chủ trì;", "diem"],
    ["đ) Các nhiệm vụ khác.", "diem"],
    ["- Như trên;", "y"],
    ["+ Tăng cường kiểm tra;", "y"],
  ])("%s → %s", (text, role) => {
    expect(detectRole(text)).toBe(role);
  });

  it.each([
    ["Chương trình công tác năm 2026 gồm các nội dung sau."],
    ["Mục tiêu của kế hoạch là nâng cao chất lượng."],
    ["Điều chỉnh kế hoạch cho phù hợp."],
    ["1.500 hộ dân được hỗ trợ trong năm."],
    ["1. 200 hộ dân được hỗ trợ."],
    ["2026 là năm bản lề."],
    ["Vì vậy, đề nghị các đơn vị thực hiện."],
    ["In đậm I. không phải đề mục."],
  ])("dễ nhầm, vẫn là thân bài: %s", (text) => {
    expect(detectRole(text)).toBe("noiDung");
  });

  it("dòng in hoa ngay sau Chương, Phần là tên chương; sau Mục là tên mục; đoạn thường sau đó là thân bài", () => {
    const roles = assignRoles(["Chương I", "QUY ĐỊNH CHUNG", "Mục 1", "PHẠM VI ÁP DỤNG", "Điều 1. Phạm vi", "Nội dung điều."]);
    expect(roles).toEqual(["chuong", "tenChuong", "muc", "tenMuc", "dieu", "noiDung"]);
    expect(assignRoles(["Chương I", "Quy định này áp dụng cho…"])).toEqual(["chuong", "noiDung"]);
  });

  it("căn cứ: dòng 'Căn cứ…;', '- Căn cứ…;' (Đảng), dòng cuối '.'; câu thân bài mở đầu bằng 'Căn cứ' vẫn là nội dung", () => {
    expect(assignRoles(["Căn cứ Luật A;", "Theo đề nghị của Văn phòng.", "Nội dung."])).toEqual(["canCu", "canCu", "noiDung"]);
    expect(assignRoles(["- Căn cứ Điều lệ Đảng;", "- Căn cứ Quy định số 399-QĐ/TW."])).toEqual(["canCu", "canCu"]);
    expect(assignRoles(["Căn cứ Kế hoạch số 12/KH-UBND, Phòng báo cáo như sau:"])).toEqual(["noiDung"]);
  });
});

describe("style theo vai trò cho từng chuẩn", () => {
  it("Chính quyền và Đảng dùng style có sẵn của chuẩn; đề mục La Mã dùng style nội dung, in đậm", () => {
    expect(roleStyle("chuong", ND30)).toEqual({ styleName: "ND30_PhanChuong" });
    expect(roleStyle("tenChuong", ND30)).toEqual({ styleName: "ND30_TieuDePhanChuong" });
    expect(roleStyle("muc", ND30)).toEqual({ styleName: "ND30_MucTieuMuc" });
    expect(roleStyle("tenMuc", HD05)).toEqual({ styleName: "HD05_TenMuc" });
    expect(roleStyle("muc", HD05)).toEqual({ styleName: "HD05_MucSo" });
    expect(roleStyle("dieu", HD05)).toEqual({ styleName: "HD05_Dieu" });
    expect(roleStyle("deMuc", ND30)).toEqual({ styleName: "ND30_NoiDung", bold: true });
    expect(roleStyle("y", ND30)).toEqual({ styleName: "ND30_NoiDung" });
    expect(roleStyle("noiDung", HD05)).toEqual({ styleName: "HD05_NoiDung" });
    // Căn cứ: Chính quyền có style căn cứ (in nghiêng); Đảng trình bày như nội dung (chữ đứng).
    expect(roleStyle("canCu", ND30)).toEqual({ styleName: "ND30_CanCu" });
    expect(roleStyle("canCu", HD05)).toEqual({ styleName: "HD05_NoiDung" });
  });

  it("kế hoạch định dạng gắn style theo vai trò, vẫn bỏ qua bảng, đoạn trống, danh sách tự động", () => {
    const plan = planFormatBody([info("Chương I"), info("QUY ĐỊNH CHUNG"), info(""), info("Điều 1. Phạm vi"), info("1. Khoản một."), { ...info("a) Điểm"), isListItem: true }, info("Nội dung.")], ND30);
    expect(plan.targets).toEqual([0, 1, 3, 4, 6]);
    expect(plan.roles.map((r) => [r.index, r.styleName])).toEqual([
      [0, "ND30_PhanChuong"],
      [1, "ND30_TieuDePhanChuong"],
      [3, "ND30_Dieu"],
      [4, "ND30_Khoan"],
      [6, "ND30_NoiDung"],
    ]);
    expect(plan.skipped.empty).toBe(1);
    expect(plan.skipped.list).toBe(1);
  });
});
