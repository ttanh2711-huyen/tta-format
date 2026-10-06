import { AGENCY_CODES, findAgencyByCode, findAgencyByName } from "./agencyCodes";

describe("HD 05: bảng tra chữ viết tắt tên cơ quan đảng (Mục I.3.1.2.b.iv)", () => {
  it("đảng uỷ và chi bộ (chữ Đ ở đầu mã)", () => {
    expect(findAgencyByName("Đảng uỷ quân sự")?.code).toBe("ĐUQS");
    expect(findAgencyByName("Quân uỷ Trung ương")?.code).toBe("QUTW");
    expect(findAgencyByName("Đảng uỷ công an")?.code).toBe("ĐUCA");
    expect(findAgencyByName("Đảng uỷ biên phòng")?.code).toBe("ĐUBP");
    expect(findAgencyByName("Các đảng uỷ khác")?.code).toBe("ĐU");
    expect(findAgencyByName("Chi bộ")?.code).toBe("CB");
  });

  it("cơ quan chuyên trách tham mưu, giúp việc", () => {
    expect(findAgencyByName("Uỷ ban Kiểm tra")?.code).toBe("UBKT");
    expect(findAgencyByName("Cơ quan uỷ ban kiểm tra")?.code).toBe("CQUBKT");
    expect(findAgencyByName("Ban tuyên giáo và dân vận")?.code).toBe("BTGDV");
    expect(findAgencyByName("Ban xây dựng đảng")?.code).toBe("BXDĐ");
  });

  it("ban chỉ đạo, tiểu ban, hội đồng và văn bản đại hội", () => {
    expect(findAgencyByName("Ban chỉ đạo")?.code).toBe("BCĐ");
    expect(findAgencyByName("Tiểu ban")?.code).toBe("TB");
    expect(findAgencyByName("Hội đồng")?.code).toBe("HĐ");
    expect(findAgencyByName("Đại hội")?.code).toBe("ĐH");
  });

  it("đơn vị sự nghiệp của Đảng", () => {
    const expected: Array<[string, string]> = [
      ["Báo Nhân Dân", "BND"],
      ["Tạp chí Cộng sản", "TCCS"],
      ["Học viện Chính trị quốc gia Hồ Chí Minh", "HVCTQGHCM"],
      ["Đài Truyền hình Việt Nam", "ĐTHVN"],
      ["Đài Tiếng nói Việt Nam", "ĐTNVN"],
      ["Thông tấn xã Việt Nam", "TTXVN"],
      ["Viện Hàn lâm Khoa học xã hội Việt Nam", "VHLKHXHVN"],
      ["Viện Hàn lâm Khoa học và Công nghệ Việt Nam", "VHLKHCNVN"],
      ["Trường chính trị", "TCT"],
      ["Trung tâm chính trị", "TTCT"],
      ["Cơ quan báo và phát thanh truyền hình", "CQBPTTH"],
    ];
    for (const [name, code] of expected) {
      expect(findAgencyByName(name)?.code).toBe(code);
    }
  });

  it("mã viết tắt không trùng nhau và chỉ gồm chữ, số, không có khoảng trắng", () => {
    const codes = AGENCY_CODES.map((a) => a.code);
    expect(new Set(codes).size).toBe(codes.length);
    for (const code of codes) {
      expect(code).toMatch(/^[\p{L}\p{N}]+$/u);
    }
  });

  it("tra cứu theo mã và theo tên không phân biệt hoa thường", () => {
    expect(findAgencyByCode("UBKT")?.name).toBe("Uỷ ban Kiểm tra");
    expect(findAgencyByName("uỷ ban kiểm tra")?.code).toBe("UBKT");
    expect(findAgencyByCode("KHONGCO")).toBeUndefined();
  });

  it("chấp nhận cả hai cách bỏ dấu \"uỷ\" và \"ủy\" khi tra theo tên", () => {
    expect(findAgencyByName("Ủy ban Kiểm tra")?.code).toBe("UBKT");
    expect(findAgencyByName("Đảng ủy công an")?.code).toBe("ĐUCA");
    expect(findAgencyByName("Đảng uỷ công an")?.code).toBe("ĐUCA");
  });
});
