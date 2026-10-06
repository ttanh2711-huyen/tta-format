import { COPY_TYPES, DOC_TYPES, findDocTypeByAbbr, findDocTypeById } from "./docTypes";

describe("NĐ 30 Điều 7: danh sách loại văn bản hành chính", () => {
  it("có đủ 29 loại, mã định danh không trùng", () => {
    expect(DOC_TYPES).toHaveLength(29);
    const ids = DOC_TYPES.map((t) => t.id);
    expect(new Set(ids).size).toBe(29);
  });

  it("27 loại có chữ viết tắt đúng như Bảng ở Phụ lục III, Mục I", () => {
    const expected: Array<[string, string]> = [
      ["Nghị quyết (cá biệt)", "NQ"],
      ["Quyết định (cá biệt)", "QĐ"],
      ["Chỉ thị", "CT"],
      ["Quy chế", "QC"],
      ["Quy định", "QyĐ"],
      ["Thông cáo", "TC"],
      ["Thông báo", "TB"],
      ["Hướng dẫn", "HD"],
      ["Chương trình", "CTr"],
      ["Kế hoạch", "KH"],
      ["Phương án", "PA"],
      ["Đề án", "ĐA"],
      ["Dự án", "DA"],
      ["Báo cáo", "BC"],
      ["Biên bản", "BB"],
      ["Tờ trình", "TTr"],
      ["Hợp đồng", "HĐ"],
      ["Công điện", "CĐ"],
      ["Bản ghi nhớ", "BGN"],
      ["Bản thỏa thuận", "BTT"],
      ["Giấy ủy quyền", "GUQ"],
      ["Giấy mời", "GM"],
      ["Giấy giới thiệu", "GGT"],
      ["Giấy nghỉ phép", "GNP"],
      ["Phiếu gửi", "PG"],
      ["Phiếu chuyển", "PC"],
      ["Phiếu báo", "PB"],
    ];
    const actual = DOC_TYPES.filter((t) => t.abbr !== null).map((t) => [t.name, t.abbr]);
    expect(actual).toEqual(expect.arrayContaining(expected));
    expect(actual).toHaveLength(27);
  });

  it("công văn và thư công không có chữ viết tắt tên loại", () => {
    expect(findDocTypeById("cong-van")).toMatchObject({ name: "Công văn", abbr: null });
    expect(findDocTypeById("thu-cong")).toMatchObject({ name: "Thư công", abbr: null });
  });

  it("chữ viết tắt không trùng nhau", () => {
    const abbrs = DOC_TYPES.map((t) => t.abbr).filter((a): a is string => a !== null);
    expect(new Set(abbrs).size).toBe(abbrs.length);
  });

  it("Mẫu 1.4 áp dụng chung cho 17 loại văn bản có tên loại", () => {
    const mau14 = DOC_TYPES.filter((t) => t.template === "Mẫu 1.4").map((t) => t.name);
    expect(mau14.sort()).toEqual(
      [
        "Chỉ thị",
        "Quy chế",
        "Quy định",
        "Thông cáo",
        "Thông báo",
        "Hướng dẫn",
        "Chương trình",
        "Kế hoạch",
        "Phương án",
        "Đề án",
        "Dự án",
        "Báo cáo",
        "Tờ trình",
        "Giấy ủy quyền",
        "Phiếu gửi",
        "Phiếu chuyển",
        "Phiếu báo",
      ].sort(),
    );
  });

  it("mỗi loại có mẫu riêng trong Phụ lục III được gán đúng số mẫu", () => {
    expect(findDocTypeById("nghi-quyet")?.template).toBe("Mẫu 1.1");
    expect(findDocTypeById("quyet-dinh")?.template).toBe("Mẫu 1.2, Mẫu 1.3");
    expect(findDocTypeById("cong-van")?.template).toBe("Mẫu 1.5");
    expect(findDocTypeById("cong-dien")?.template).toBe("Mẫu 1.6");
    expect(findDocTypeById("giay-moi")?.template).toBe("Mẫu 1.7");
    expect(findDocTypeById("giay-gioi-thieu")?.template).toBe("Mẫu 1.8");
    expect(findDocTypeById("bien-ban")?.template).toBe("Mẫu 1.9");
    expect(findDocTypeById("giay-nghi-phep")?.template).toBe("Mẫu 1.10");
  });

  it("hợp đồng, bản ghi nhớ, bản thỏa thuận và thư công chưa có mẫu trong Phụ lục III", () => {
    for (const id of ["hop-dong", "ban-ghi-nho", "ban-thoa-thuan", "thu-cong"]) {
      expect(findDocTypeById(id)).toMatchObject({ template: null, layout: "chua-co-mau" });
    }
  });

  it("tra cứu theo chữ viết tắt", () => {
    expect(findDocTypeByAbbr("TTr")?.name).toBe("Tờ trình");
    expect(findDocTypeByAbbr("QyĐ")?.name).toBe("Quy định");
    expect(findDocTypeByAbbr("không-có")).toBeUndefined();
  });

  it("bản sao văn bản gồm SY, TrS, SL", () => {
    expect(COPY_TYPES.map((c) => c.abbr)).toEqual(["SY", "TrS", "SL"]);
  });
});
