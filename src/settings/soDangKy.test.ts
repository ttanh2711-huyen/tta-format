import AdmZip from "adm-zip";
import { buildTableDocx } from "../checker/exportDocx";
import { buildXlsx, columnName } from "../ooxml/xlsx";
import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { addEntry, entryFromDocument, listEntries, MAX_ENTRIES, registerRows, REGISTER_HEADER, registerYears, removeEntry } from "./soDangKy";
import { exportConfig, importConfig } from "./exportImport";
import { memoryStore } from "./storage";

const at = (iso: string) => new Date(iso);
const sample = { soKyHieu: "Số: 05/CCTL-VP", noiNgay: "Cần Thơ, ngày 05 tháng 01 năm 2026", loai: "Công văn", trichYeu: "V/v cử công chức dự tập huấn", noiNhan: "Sở Nông nghiệp và Môi trường; Lưu: VT", nguoiKy: "Mai Hiếu Hiền" };

describe("sổ đăng ký văn bản đi", () => {
  it("ghi, liệt kê mới nhất trước, theo chuẩn riêng, xoá được", () => {
    const store = memoryStore();
    const a = addEntry(store, ND30, { ...sample, year: 2026 }, at("2026-01-05T08:00:00Z"));
    const b = addEntry(store, ND30, { ...sample, year: 2026, soKyHieu: "Số: 06/CCTL-VP" }, at("2026-01-06T08:00:00Z"));
    addEntry(store, HD05, { ...sample, year: 2026 }, at("2026-01-07T08:00:00Z"));
    expect(listEntries(store, ND30).map((e) => e.soKyHieu)).toEqual(["Số: 06/CCTL-VP", "Số: 05/CCTL-VP"]);
    expect(listEntries(store, HD05)).toHaveLength(1);
    removeEntry(store, ND30, a.id);
    expect(listEntries(store, ND30).map((e) => e.id)).toEqual([b.id]);
  });

  it("lọc theo năm, liệt kê các năm có trong sổ", () => {
    const store = memoryStore();
    addEntry(store, ND30, { ...sample, year: 2025 }, at("2025-12-30T08:00:00Z"));
    addEntry(store, ND30, { ...sample, year: 2026 }, at("2026-01-02T08:00:00Z"));
    expect(listEntries(store, ND30, 2025)).toHaveLength(1);
    expect(registerYears(store, ND30)).toEqual([2026, 2025]);
  });

  it("giữ tối đa MAX_ENTRIES mục, bỏ mục cũ nhất; dữ liệu hỏng thì coi như sổ trống", () => {
    const store = memoryStore();
    for (let i = 0; i < MAX_ENTRIES + 3; i += 1) addEntry(store, ND30, { ...sample, year: 2026, soKyHieu: `Số: ${i}` }, at(`2026-01-01T00:00:${String(i % 60).padStart(2, "0")}Z`));
    expect(listEntries(store, ND30)).toHaveLength(MAX_ENTRIES);
    store.set(`nd30.dk.${ND30.id}`, "không phải JSON");
    expect(listEntries(store, ND30)).toEqual([]);
  });

  it("dựng mục từ dữ liệu văn bản vừa tạo: gộp dòng nơi nhận, bỏ khoảng trắng thừa, năm theo ngày văn bản", () => {
    const entry = entryFromDocument({ typeName: "Công văn", so: "Số:  05/CCTL-VP ", placeDate: "Cần Thơ, ngày 05 tháng 01 năm 2026", subject: " cử công chức ", recipients: ["Sở Nông nghiệp", " Như trên "], signer: "Mai Hiếu Hiền" }, at("2025-12-31T00:00:00Z"));
    expect(entry).toMatchObject({ year: 2026, soKyHieu: "Số: 05/CCTL-VP", trichYeu: "cử công chức", noiNhan: "Sở Nông nghiệp; Như trên", nguoiKy: "Mai Hiếu Hiền" });
    expect(entryFromDocument({ typeName: "", so: "", placeDate: "Cần Thơ, ngày tháng năm", subject: "", recipients: [], signer: "" }, at("2026-03-01T00:00:00Z")).year).toBe(2026);
  });

  it("bảng xuất có tiêu đề cột và số thứ tự theo thứ tự ghi sổ (cũ trước)", () => {
    const store = memoryStore();
    addEntry(store, ND30, { ...sample, year: 2026, soKyHieu: "Số: 01" }, at("2026-01-01T00:00:00Z"));
    addEntry(store, ND30, { ...sample, year: 2026, soKyHieu: "Số: 02" }, at("2026-01-02T00:00:00Z"));
    const rows = registerRows(listEntries(store, ND30));
    expect(rows[0]).toEqual([...REGISTER_HEADER]);
    expect(rows.slice(1).map((r) => [r[0], r[1]])).toEqual([["1", "Số: 01"], ["2", "Số: 02"]]);
  });
});

describe("tệp Excel và bảng Word của sổ", () => {
  const rows = [["STT", "Trích yếu"], ["1", "Báo cáo <quý> & năm"]];

  it("tên cột Excel A, Z, AA", () => {
    expect([columnName(0), columnName(25), columnName(26)]).toEqual(["A", "Z", "AA"]);
  });

  it(".xlsx là zip hợp lệ, có trang tính với chữ đã thoát ký tự XML", () => {
    const zip = new AdmZip(Buffer.from(buildXlsx("Sổ văn bản đi", rows)));
    expect(zip.getEntries().map((e) => e.entryName)).toEqual(expect.arrayContaining(["[Content_Types].xml", "xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/styles.xml"]));
    const sheet = zip.readAsText("xl/worksheets/sheet1.xml");
    expect(sheet).toContain("Báo cáo &lt;quý&gt; &amp; năm");
    expect(sheet).toContain('r="B2"');
    expect(zip.readAsText("xl/workbook.xml")).toContain('name="Sổ văn bản đi"');
  });

  it("bảng Word: tiêu đề và đủ ô", () => {
    const xml = new AdmZip(Buffer.from(buildTableDocx("SỔ ĐĂNG KÝ VĂN BẢN ĐI", rows))).readAsText("word/document.xml");
    expect(xml).toContain("SỔ ĐĂNG KÝ VĂN BẢN ĐI");
    expect(xml.match(/<w:tc>/g)).toHaveLength(4);
    expect(xml).toContain("Báo cáo &lt;quý&gt; &amp; năm");
  });
});

describe("sổ đăng ký trong xuất, nhập cấu hình", () => {
  it("xuất rồi nhập sang máy khác giữ đủ dòng; nhập lại không nhân đôi", () => {
    const from = memoryStore();
    addEntry(from, ND30, { ...sample, year: 2026 }, at("2026-01-05T08:00:00Z"));
    addEntry(from, HD05, { ...sample, year: 2026 }, at("2026-01-06T08:00:00Z"));
    const text = exportConfig(from);
    const to = memoryStore();
    expect(importConfig(to, text).registerEntries).toBe(2);
    expect(importConfig(to, text).registerEntries).toBe(0);
    expect(listEntries(to, ND30)).toEqual(listEntries(from, ND30));
  });
});
