import { DOC_TYPES, findDocTypeById, findDocTypesByAbbr } from "./docTypes";

describe("Quy định 399-QĐ/TW: thể loại văn bản của Đảng", () => {
  it("Điều 5 có 25 thể loại, Điều 6 có 8 loại giấy tờ hành chính, mã định danh không trùng", () => {
    expect(DOC_TYPES.filter((t) => t.layout !== "giay-to")).toHaveLength(25);
    expect(DOC_TYPES.filter((t) => t.layout === "giay-to")).toHaveLength(8);
    expect(new Set(DOC_TYPES.map((t) => t.id)).size).toBe(33);
  });

  it("danh sách tên đúng như Điều 5 và Điều 6", () => {
    const dieu5 = [
      "Cương lĩnh chính trị",
      "Điều lệ Đảng",
      "Chiến lược",
      "Nghị quyết",
      "Quyết định",
      "Chỉ thị",
      "Kết luận",
      "Quy chế",
      "Quy định",
      "Thông tri",
      "Hướng dẫn",
      "Thông báo",
      "Thông cáo",
      "Tuyên bố",
      "Lời kêu gọi",
      "Báo cáo",
      "Kế hoạch",
      "Quy hoạch",
      "Chương trình",
      "Đề án",
      "Phương án",
      "Dự án",
      "Tờ trình",
      "Công văn",
      "Biên bản",
    ];
    const dieu6 = [
      "Giấy giới thiệu",
      "Giấy chứng nhận",
      "Giấy đi đường",
      "Giấy nghỉ phép",
      "Giấy mời",
      "Phiếu chuyển",
      "Phiếu gửi",
      "Thư công",
    ];
    expect(DOC_TYPES.map((t) => t.name)).toEqual([...dieu5, ...dieu6]);
  });

  it("chữ viết tắt nêu rõ trong Hướng dẫn 05 (Mục I.3.1.2 và ví dụ)", () => {
    const explicit = DOC_TYPES.filter((t) => t.abbrSource === "explicit").map((t) => [t.name, t.abbr]);
    expect(explicit).toEqual(
      expect.arrayContaining([
        ["Nghị quyết", "NQ"],
        ["Chỉ thị", "CT"],
        ["Kết luận", "KL"],
        ["Quy chế", "QC"],
        ["Báo cáo", "BC"],
        ["Quyết định", "QĐ"],
        ["Quy định", "QĐ"],
        ["Chương trình", "CTr"],
        ["Thông tri", "TT"],
        ["Tờ trình", "TTr"],
        ["Công văn", "CV"],
        ["Hướng dẫn", "HD"],
        ["Biên bản", "BB"],
      ]),
    );
    expect(explicit).toHaveLength(13);
  });

  it("quyết định và quy định cùng dùng QĐ", () => {
    expect(findDocTypesByAbbr("QĐ").map((t) => t.id).sort()).toEqual(["quy-dinh", "quyet-dinh"]);
  });

  it("chữ viết tắt suy ra theo quy tắc chữ cái đầu được đánh dấu để người dùng kiểm tra", () => {
    for (const id of ["thong-bao", "thong-cao", "ke-hoach", "quy-hoach", "de-an", "phuong-an", "du-an"]) {
      expect(findDocTypeById(id)).toMatchObject({ abbrSource: "derived" });
    }
    expect(findDocTypeById("thong-bao")?.abbr).toBe("TB");
    expect(findDocTypeById("de-an")?.abbr).toBe("ĐA");
  });

  it("tuyên bố trùng chữ cái đầu với thông báo nên chưa gán chữ viết tắt", () => {
    expect(findDocTypeById("tuyen-bo")).toMatchObject({ abbr: null });
    expect(findDocTypeById("tuyen-bo")?.note).toMatch(/TB/);
  });

  it("các văn bản không có số, ký hiệu thông thường không có chữ viết tắt", () => {
    for (const id of ["cuong-linh-chinh-tri", "dieu-le-dang", "chien-luoc", "loi-keu-goi", "thu-cong"]) {
      expect(findDocTypeById(id)?.abbr).toBeNull();
    }
  });

  it("công văn dùng bố cục công văn, thể loại khác trong Điều 5 dùng bố cục có tên loại", () => {
    expect(findDocTypeById("cong-van")?.layout).toBe("cong-van");
    expect(findDocTypeById("bao-cao")?.layout).toBe("co-ten-loai");
  });

  it("tra cứu chữ viết tắt không tồn tại trả về danh sách rỗng", () => {
    expect(findDocTypesByAbbr("không-có")).toEqual([]);
  });
});
