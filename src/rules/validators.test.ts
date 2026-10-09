import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { placeDateText, soText } from "./validators";

const GAP = "      ";

describe("soText: số và ký hiệu", () => {
  it("số nhỏ hơn 10 thêm số 0, ký hiệu nối bằng - và /", () => {
    expect(soText(ND30, { so: 5, typeAbbr: "QĐ", agencyAbbrs: ["UBND"] })).toEqual({ ok: true, text: "Số: 05/QĐ-UBND" });
    expect(soText(ND30, { so: 15, typeAbbr: null, agencyAbbrs: ["BNV", "VP"] })).toEqual({ ok: true, text: "Số: 15/BNV-VP" });
    expect(soText(ND30, { so: 10, typeAbbr: null, agencyAbbrs: ["A"] })).toEqual({ ok: true, text: "Số: 10/A" });
  });

  it("để trống số chờ vào sổ: chừa chỗ trống đúng vị trí", () => {
    expect(soText(ND30, { so: null, typeAbbr: null, agencyAbbrs: ["CCTL", "ĐĐPCTT"] })).toEqual({ ok: true, text: `Số: ${GAP}/CCTL-ĐĐPCTT` });
  });

  it("chuẩn Đảng: không có dấu hai chấm, ký hiệu loại trước dấu /", () => {
    expect(soText(HD05, { so: 5, typeAbbr: "HD", agencyAbbrs: ["VPTW"] })).toEqual({ ok: true, text: "Số 05-HD/VPTW" });
    expect(soText(HD05, { so: null, typeAbbr: "HD", agencyAbbrs: ["VPTW"] })).toEqual({ ok: true, text: `Số ${GAP}-HD/VPTW` });
  });

  it("chuẩn Đảng: các ví dụ trong hướng dẫn, kể cả liên cơ quan và văn bản đại hội", () => {
    expect(soText(HD05, { so: 127, typeAbbr: "QĐ", agencyAbbrs: ["TW"] })).toEqual({ ok: true, text: "Số 127-QĐ/TW" });
    expect(soText(HD05, { so: 5, typeAbbr: "QC", agencyAbbrs: ["BTCTU", "BNCTU"] })).toEqual({ ok: true, text: "Số 05-QC/BTCTU-BNCTU" });
    expect(soText(HD05, { so: 1, typeAbbr: "BC", agencyAbbrs: ["ĐH"] })).toEqual({ ok: true, text: "Số 01-BC/ĐH" });
  });

  it.each([0, 1.5, -3])("số %p không hợp lệ bị từ chối", (so) => {
    expect(soText(ND30, { so, typeAbbr: null, agencyAbbrs: ["A"] })).toEqual({ ok: false, error: "Số văn bản phải là số nguyên dương." });
  });

  it("ký hiệu có khoảng trắng hoặc dấu phân cách bị từ chối", () => {
    expect(soText(ND30, { so: 1, typeAbbr: "QĐ", agencyAbbrs: ["UB ND"] })).toMatchObject({ ok: false, error: expect.stringContaining("không có khoảng trắng") });
    expect(soText(ND30, { so: 1, typeAbbr: "Q-Đ", agencyAbbrs: ["UBND"] })).toMatchObject({ ok: false });
    expect(soText(ND30, { so: 1, typeAbbr: null, agencyAbbrs: [] })).toMatchObject({ ok: false, error: expect.stringContaining("ít nhất một") });
  });

  it("chuẩn Đảng thiếu ký hiệu loại bị từ chối", () => {
    expect(soText(HD05, { so: 1, typeAbbr: null, agencyAbbrs: ["TW"] })).toMatchObject({ ok: false, error: expect.stringContaining("tên loại") });
  });

  it("số trống vẫn kiểm tra ký hiệu", () => {
    expect(soText(ND30, { so: null, typeAbbr: null, agencyAbbrs: [] })).toMatchObject({ ok: false });
  });
});

describe("placeDateText: địa danh và ngày tháng", () => {
  it("đủ ngày tháng năm: ngày < 10 và tháng 1, 2 thêm số 0", () => {
    expect(placeDateText({ place: "Hà Nội", day: 5, month: 1, year: 2020 })).toEqual({ ok: true, text: "Hà Nội, ngày 05 tháng 01 năm 2020" });
    expect(placeDateText({ place: "Hà Nội", day: 5, month: 3, year: 2020 })).toEqual({ ok: true, text: "Hà Nội, ngày 05 tháng 3 năm 2020" });
    expect(placeDateText({ place: "Hà Nội", day: 15, month: 12, year: 2020 })).toEqual({ ok: true, text: "Hà Nội, ngày 15 tháng 12 năm 2020" });
  });

  it("để trống ngày chờ ký như file mẫu", () => {
    expect(placeDateText({ place: "Cần Thơ", day: null, month: 10, year: 2026 })).toEqual({ ok: true, text: `Cần Thơ, ngày ${GAP} tháng 10 năm 2026` });
    expect(placeDateText({ place: "Cần Thơ", day: null, month: null, year: 2026 })).toEqual({ ok: true, text: `Cần Thơ, ngày ${GAP} tháng ${GAP} năm 2026` });
    expect(placeDateText({ place: "Cần Thơ", day: null, month: null, year: null })).toEqual({ ok: true, text: `Cần Thơ, ngày ${GAP} tháng ${GAP} năm ${GAP}` });
  });

  it("chỉ có ngày hoặc chỉ có tháng: vẫn thêm số 0", () => {
    expect(placeDateText({ place: "A", day: 5, month: null, year: 2026 })).toMatchObject({ text: `A, ngày 05 tháng ${GAP} năm 2026` });
    expect(placeDateText({ place: "A", day: null, month: 2, year: 2026 })).toMatchObject({ text: `A, ngày ${GAP} tháng 02 năm 2026` });
  });

  it("ngày không tồn tại bị từ chối: 31/2, 30/2, 31/4, 29/2 năm không nhuận", () => {
    expect(placeDateText({ place: "A", day: 31, month: 2, year: 2024 })).toMatchObject({ ok: false });
    expect(placeDateText({ place: "A", day: 30, month: 2, year: null })).toMatchObject({ ok: false });
    expect(placeDateText({ place: "A", day: 31, month: 4, year: 2024 })).toMatchObject({ ok: false });
    expect(placeDateText({ place: "A", day: 29, month: 2, year: 2025 })).toEqual({ ok: false, error: "Tháng 2 năm 2025 không có ngày 29." });
  });

  it("29/2 hợp lệ năm nhuận, hoặc khi chưa biết năm", () => {
    expect(placeDateText({ place: "A", day: 29, month: 2, year: 2024 })).toMatchObject({ ok: true });
    expect(placeDateText({ place: "A", day: 29, month: 2, year: null })).toMatchObject({ ok: true });
  });

  it.each([
    [{ day: 0, month: 1, year: 2020 }, "Ngày phải từ 1 đến 31."],
    [{ day: 32, month: 1, year: 2020 }, "Ngày phải từ 1 đến 31."],
    [{ day: 1, month: 13, year: 2020 }, "Tháng phải từ 1 đến 12."],
    [{ day: 1, month: 0, year: 2020 }, "Tháng phải từ 1 đến 12."],
    [{ day: 1, month: 1, year: 20 }, "Năm phải là số nguyên gồm 4 chữ số."],
    [{ day: 1.5, month: 1, year: 2020 }, "Ngày phải từ 1 đến 31."],
  ])("giá trị ngoài khoảng bị từ chối %#", (input, error) => {
    expect(placeDateText({ place: "A", ...input })).toEqual({ ok: false, error });
  });

  it("cấp hành chính ghi trước địa danh được giữ nguyên", () => {
    expect(placeDateText({ place: "TP. Huế", day: 3, month: 2, year: 2026 })).toEqual({ ok: true, text: "TP. Huế, ngày 03 tháng 02 năm 2026" });
    expect(placeDateText({ place: "Phường Ngô Quyền", day: 20, month: 10, year: 2025 })).toMatchObject({ ok: true, text: "Phường Ngô Quyền, ngày 20 tháng 10 năm 2025" });
  });

  it("địa danh bắt buộc, khoảng trắng thừa được gọn", () => {
    expect(placeDateText({ place: "   ", day: 1, month: 1, year: 2020 })).toEqual({ ok: false, error: "Hãy nhập địa danh." });
    expect(placeDateText({ place: "  Cần   Thơ ", day: 1, month: 1, year: 2020 })).toMatchObject({ ok: true, text: "Cần Thơ, ngày 01 tháng 01 năm 2020" });
  });
});
