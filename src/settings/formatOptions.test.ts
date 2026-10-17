import { buildStyleSpecs } from "../ooxml/styleSpecs";
import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import type { StandardProfile } from "../standards/types";
import {
  buildRegistry,
  checkValue,
  defaultValues,
  initialState,
  reconcile,
  resetGroup,
  resetOption,
  setOption,
  toPageSetupSettings,
  toStyleOptions,
  type FormatRegistry,
  type FormatState,
  type OptionDef,
} from "./formatOptions";

function option(registry: FormatRegistry, id: string): OptionDef {
  const found = registry.options.find((o) => o.id === id);
  if (!found) {
    throw new Error(`Không có tuỳ chọn ${id}; có: ${registry.options.map((o) => o.id).join(", ")}`);
  }
  return found;
}

function applied(result: ReturnType<typeof setOption>): FormatState {
  if (result.status !== "applied") {
    throw new Error(`Kỳ vọng applied, nhận ${result.status}`);
  }
  return result.state;
}

const nd = buildRegistry(ND30);
const hd = buildRegistry(HD05);

describe("registry: cấu trúc", () => {
  it.each([ND30, HD05])("$id: mã tuỳ chọn không trùng, thuộc nhóm đã khai báo, có nhãn tiếng Việt", (profile) => {
    const registry = buildRegistry(profile);
    const ids = registry.options.map((o) => o.id);
    expect(new Set(ids).size).toBe(ids.length);
    const groups = new Set(registry.groups.map((g) => g.id));
    for (const o of registry.options) {
      expect(groups.has(o.group)).toBe(true);
      expect(o.label.length).toBeGreaterThan(0);
    }
    expect(registry.groups.map((g) => g.id)).toEqual(["page", "font", "paragraph", "header", "footer", "table", "pageNumber", "behavior"]);
  });

  it.each([ND30, HD05])("$id: mặc định của mọi tuỳ chọn hợp lệ với chính chuẩn (không cần xác nhận)", (profile) => {
    const registry = buildRegistry(profile);
    for (const o of registry.options) {
      expect({ id: o.id, check: checkValue(o, o.default) }).toEqual({ id: o.id, check: { status: "ok" } });
    }
  });

  it.each([ND30, HD05])("$id: mọi giá trị chọn nhanh đều hợp lệ trừ giá trị đánh dấu ngoại lệ", (profile) => {
    for (const o of buildRegistry(profile).options) {
      for (const choice of o.choices) {
        const result = checkValue(o, choice.value);
        const fine = result.status === "ok" || choice.exception !== undefined;
        expect({ id: o.id, value: choice.value, fine }).toEqual({ id: o.id, value: choice.value, fine: true });
      }
    }
  });

  it("chuẩn Chính quyền không có tuỳ chọn in hai mặt, chuẩn Đảng có", () => {
    expect(nd.options.some((o) => o.id === "duplex")).toBe(false);
    expect(hd.options.some((o) => o.id === "duplex")).toBe(true);
  });

  it("mỗi chuẩn có registry riêng: khoảng giá trị lấy từ chính chuẩn đó", () => {
    expect(option(nd, "lineSpacing").unit).toBe("dòng");
    expect(option(hd, "lineSpacing").unit).toBe("pt");
    expect(option(hd, "lineSpacing").choices.map((c) => c.value)).toEqual([18, 19, 20, 21, 22]);
    expect(option(nd, "lineSpacing").choices.map((c) => c.value)).toEqual([1, 1.17, 1.5]);
    expect(option(nd, "fontSize").choices.map((c) => c.value)).toEqual([13, 14]);
    expect(option(hd, "fontSize").choices.map((c) => c.value)).toEqual([14, 15]);
  });

  it("có đủ các nhóm tuỳ chọn theo bảng trong kế hoạch", () => {
    const ids = (registry: FormatRegistry, group: string) => registry.options.filter((o) => o.group === group).map((o) => o.id);
    expect(ids(nd, "page")).toEqual(["marginTop", "marginBottom", "marginLeft", "marginRight", "headerDistance", "footerDistance"]);
    expect(ids(hd, "page")).toContain("duplex");
    expect(ids(nd, "font")).toEqual(["fontSize"]);
    expect(ids(nd, "paragraph")).toEqual(["alignment", "firstLineIndent", "spacing", "lineSpacing"]);
    expect(ids(nd, "table")).toEqual(["tableFontSize", "tableSpacing", "tableBorder", "tableHeaderBold"]);
    expect(ids(nd, "pageNumber")).toEqual(["pageNumberSize", "pageNumberHideFirst", "pageNumberAppendixSeparate"]);
    expect(ids(nd, "behavior")).toEqual(["localFormatting"]);
    expect(ids(nd, "header")).toContain("size.quocHieu");
    expect(ids(hd, "header")).toContain("size.theLoai");
    expect(ids(nd, "footer")).toEqual(expect.arrayContaining(["signGap", "size.noiNhanDanhSach"]));
  });

  it("tác dụng: trang, chữ, đoạn, bảng (cỡ, giãn) áp ngay; đầu, cuối, số trang áp khi tạo văn bản", () => {
    expect(option(nd, "marginLeft").effect).toBe("now");
    expect(option(nd, "fontSize").effect).toBe("now");
    expect(option(nd, "tableFontSize").effect).toBe("now");
    expect(option(nd, "tableBorder").effect).toBe("onCreate");
    expect(option(nd, "size.quocHieu").effect).toBe("onCreate");
    expect(option(nd, "pageNumberSize").effect).toBe("onCreate");
    expect(option(nd, "localFormatting").effect).toBe("behavior");
  });
});

describe("mặc định khớp bảng quy tắc và file mẫu", () => {
  it("chuẩn Chính quyền: 20/20/30/20, header 12,5, footer 5,3, 14 pt, thụt 12,7 mm, giãn 6 pt, dòng 1,17", () => {
    expect(defaultValues(nd)).toMatchObject({
      marginTop: 20,
      marginBottom: 20,
      marginLeft: 30,
      marginRight: 20,
      headerDistance: 12.5,
      footerDistance: 5.3,
      fontSize: 14,
      alignment: "Justified",
      firstLineIndent: 12.7,
      spacing: 6,
      lineSpacing: 1.17,
      tableFontSize: 14,
      tableSpacing: 3,
      localFormatting: "keep",
    });
  });

  it("chuẩn Đảng: 20/20/30/15, Exactly 20 pt, thụt 10 mm, cỡ khối đầu theo HD 05", () => {
    expect(defaultValues(hd)).toMatchObject({
      marginRight: 15,
      lineSpacing: 20,
      firstLineIndent: 10,
      fontSize: 14,
      duplex: false,
      "size.tieuDe": 15,
      "size.theLoai": 16,
    });
  });

  it("cỡ khối đầu của chuẩn Chính quyền theo bảng kế hoạch", () => {
    expect(defaultValues(nd)).toMatchObject({
      "size.quocHieu": 13,
      "size.tieuNgu": 14,
      "size.diaDanhNgayThang": 13,
      "size.noiNhanTieuDe": 12,
      "size.noiNhanDanhSach": 11,
    });
  });
});

describe("checkValue", () => {
  it("giá trị trong khoảng của chuẩn thì hợp lệ", () => {
    expect(checkValue(option(nd, "marginLeft"), 32)).toEqual({ status: "ok" });
    expect(checkValue(option(nd, "lineSpacing"), 1.3)).toEqual({ status: "ok" });
  });

  it("ngoài khoảng của chuẩn thì cần xác nhận, kèm lý do tiếng Việt nêu cả chuẩn", () => {
    const result = checkValue(option(nd, "fontSize"), 16);
    expect(result).toMatchObject({ status: "exception", reason: expect.stringContaining("Chính quyền") });
    expect(checkValue(option(nd, "marginLeft"), 40)).toMatchObject({ status: "exception" });
  });

  it("giá trị hợp lệ của chuẩn này không hợp lệ ở chuẩn kia", () => {
    expect(checkValue(option(nd, "marginRight"), 20)).toEqual({ status: "ok" });
    expect(checkValue(option(hd, "marginRight"), 20)).toMatchObject({ status: "exception" });
    expect(checkValue(option(hd, "lineSpacing"), 22)).toEqual({ status: "ok" });
    // Exactly 20 pt vô nghĩa với giãn dòng bội số: bị từ chối hẳn, không cho xác nhận ngoại lệ.
    expect(checkValue(option(nd, "lineSpacing"), 20)).toMatchObject({ status: "invalid" });
    expect(checkValue(option(nd, "lineSpacing"), 2)).toMatchObject({ status: "exception" });
  });

  it("thụt dòng đầu chỉ nhận đúng các giá trị chuẩn cho phép", () => {
    expect(checkValue(option(nd, "firstLineIndent"), 10)).toEqual({ status: "ok" });
    expect(checkValue(option(nd, "firstLineIndent"), 11)).toMatchObject({ status: "exception" });
  });

  it("giãn đoạn dưới mức tối thiểu cần xác nhận", () => {
    expect(checkValue(option(nd, "spacing"), 3)).toMatchObject({ status: "exception", reason: expect.stringContaining("6") });
  });

  it("không phải số, vô cực, âm, ngoài giới hạn cứng thì bị từ chối hẳn", () => {
    const margin = option(nd, "marginLeft");
    expect(checkValue(margin, Number.NaN)).toMatchObject({ status: "invalid" });
    expect(checkValue(margin, Number.POSITIVE_INFINITY)).toMatchObject({ status: "invalid" });
    expect(checkValue(margin, -5)).toMatchObject({ status: "invalid" });
    expect(checkValue(margin, 500)).toMatchObject({ status: "invalid" });
    expect(checkValue(margin, "30")).toMatchObject({ status: "invalid" });
  });

  it("lựa chọn: giá trị lạ bị từ chối, căn trái là ngoại lệ cần xác nhận", () => {
    const alignment = option(nd, "alignment");
    expect(checkValue(alignment, "Justified")).toEqual({ status: "ok" });
    expect(checkValue(alignment, "Left")).toMatchObject({ status: "exception" });
    expect(checkValue(alignment, "Centered")).toMatchObject({ status: "invalid" });
  });

  it("tuỳ chọn bật tắt chỉ nhận true hoặc false", () => {
    expect(checkValue(option(nd, "tableBorder"), false)).toEqual({ status: "ok" });
    expect(checkValue(option(nd, "tableBorder"), 1)).toMatchObject({ status: "invalid" });
  });
});

describe("setOption, resetOption, resetGroup", () => {
  it("giá trị hợp lệ được áp dụng và không đánh dấu ngoại lệ", () => {
    const state = applied(setOption(nd, initialState(nd), "fontSize", 13));
    expect(state.values.fontSize).toBe(13);
    expect(state.exceptions).toEqual([]);
  });

  it("giá trị ngoài khoảng chưa xác nhận thì không đổi gì và trả lý do", () => {
    const state = initialState(nd);
    const result = setOption(nd, state, "fontSize", 16);
    expect(result).toMatchObject({ status: "needsConfirm", reason: expect.any(String) });
    expect(state.values.fontSize).toBe(14);
  });

  it("đã xác nhận thì áp dụng và đánh dấu ngoại lệ chủ ý", () => {
    const state = applied(setOption(nd, initialState(nd), "fontSize", 16, { confirmed: true }));
    expect(state.values.fontSize).toBe(16);
    expect(state.exceptions).toEqual(["fontSize"]);
  });

  it("đổi về giá trị hợp lệ thì bỏ đánh dấu ngoại lệ", () => {
    const first = applied(setOption(nd, initialState(nd), "fontSize", 16, { confirmed: true }));
    const second = applied(setOption(nd, first, "fontSize", 14));
    expect(second.exceptions).toEqual([]);
  });

  it("giá trị không hợp lệ bị từ chối dù đã xác nhận; tuỳ chọn lạ bị từ chối", () => {
    expect(setOption(nd, initialState(nd), "marginLeft", Number.NaN, { confirmed: true }).status).toBe("rejected");
    expect(setOption(nd, initialState(nd), "khongCo", 1).status).toBe("rejected");
  });

  it("resetOption đưa một tuỳ chọn về mặc định và bỏ ngoại lệ", () => {
    const set = applied(setOption(nd, initialState(nd), "fontSize", 16, { confirmed: true }));
    const reset = resetOption(nd, set, "fontSize");
    expect(reset.values.fontSize).toBe(14);
    expect(reset.exceptions).toEqual([]);
  });

  it("resetGroup chỉ đặt lại các tuỳ chọn của nhóm đó", () => {
    let state = initialState(nd);
    state = applied(setOption(nd, state, "fontSize", 13));
    state = applied(setOption(nd, state, "lineSpacing", 1.5));
    const reset = resetGroup(nd, state, "font");
    expect(reset.values.fontSize).toBe(14);
    expect(reset.values.lineSpacing).toBe(1.5);
  });
});

describe("reconcile: giá trị đã lưu so với registry hiện hành", () => {
  it("giữ giá trị hợp lệ, điền mặc định cho tuỳ chọn thiếu", () => {
    const { state, ignored } = reconcile(nd, { fontSize: 13 }, []);
    expect(state.values.fontSize).toBe(13);
    expect(state.values.marginLeft).toBe(30);
    expect(ignored).toEqual([]);
  });

  it("bỏ qua tuỳ chọn không còn tồn tại và giá trị sai, có thông báo", () => {
    const { state, ignored } = reconcile(nd, { cuCu: 5, fontSize: "to" }, []);
    expect(state.values.fontSize).toBe(14);
    expect(state.values).not.toHaveProperty("cuCu");
    expect(ignored).toHaveLength(2);
    expect(ignored.join(" ")).toContain("cuCu");
  });

  it("giá trị ngoài khoảng chỉ được giữ nếu đã đánh dấu ngoại lệ chủ ý", () => {
    const kept = reconcile(nd, { fontSize: 16 }, ["fontSize"]);
    expect(kept.state.values.fontSize).toBe(16);
    expect(kept.state.exceptions).toEqual(["fontSize"]);
    const dropped = reconcile(nd, { fontSize: 16 }, []);
    expect(dropped.state.values.fontSize).toBe(14);
    expect(dropped.ignored).toHaveLength(1);
  });

  it("đánh dấu ngoại lệ thừa (giá trị đã nằm trong khoảng) bị bỏ", () => {
    expect(reconcile(nd, { fontSize: 14 }, ["fontSize"]).state.exceptions).toEqual([]);
  });
});

describe("ánh xạ sang style và trang", () => {
  it("toStyleOptions: mặc định cho ra đúng bộ style mặc định của chuẩn", () => {
    expect(buildStyleSpecs(ND30, toStyleOptions(ND30, defaultValues(nd)))).toEqual(buildStyleSpecs(ND30));
    expect(buildStyleSpecs(HD05, toStyleOptions(HD05, defaultValues(hd)))).toEqual(buildStyleSpecs(HD05));
  });

  it("toStyleOptions: đổi cỡ chữ, giãn dòng, căn lề, bảng", () => {
    const values = { ...defaultValues(nd), fontSize: 13, lineSpacing: 1.5, spacing: 8, alignment: "Left", tableFontSize: 13, tableSpacing: 0 };
    expect(toStyleOptions(ND30, values)).toMatchObject({
      fontSizePt: 13,
      spaceBeforePt: 8,
      spaceAfterPt: 8,
      alignment: "Left",
      lineSpacing: { kind: "multiple", lines: 1.5 },
      table: { fontSizePt: 13, spaceBeforePt: 0, spaceAfterPt: 0 },
    });
  });

  it("toStyleOptions: chuẩn Đảng dùng Exactly", () => {
    expect(toStyleOptions(HD05, { ...defaultValues(hd), lineSpacing: 22 }).lineSpacing).toEqual({ kind: "exactly", pt: 22 });
  });

  it("toPageSetupSettings: lề, khoảng cách header footer, lề đối xứng", () => {
    expect(toPageSetupSettings({ ...defaultValues(hd), duplex: true })).toEqual({
      paperSize: "A4",
      orientation: "Portrait",
      marginsMm: { top: 20, bottom: 20, left: 30, right: 15 },
      headerDistanceMm: 10,
      footerDistanceMm: 12.5,
      mirrorMargins: true,
    });
    expect(toPageSetupSettings(defaultValues(nd)).mirrorMargins).toBe(false);
  });
});

describe("mã dùng chung không rẽ nhánh theo mã chuẩn", () => {
  it("registry dựng được cho một chuẩn có khoảng khác mà không cần sửa mã", () => {
    const fake: StandardProfile = {
      ...ND30,
      page: { ...ND30.page, marginsMm: { ...ND30.page.marginsMm, left: { min: 25, max: 25, source: "thử" } } },
      pageDefaults: { ...ND30.pageDefaults, marginsMm: { ...ND30.pageDefaults.marginsMm, left: 25 } },
    };
    const registry = buildRegistry(fake);
    expect(checkValue(option(registry, "marginLeft"), 25)).toEqual({ status: "ok" });
    expect(checkValue(option(registry, "marginLeft"), 30)).toMatchObject({ status: "exception" });
  });
});
