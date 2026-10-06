import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { isMarginValid } from "../standards/validate";
import { mmToPt } from "../standards/units";
import {
  findPageSetupMismatches,
  parseMm,
  planPageSetup,
  summarizeApply,
  type PageSetupSettings,
} from "./pageSetupPlan";

function settingsOf(plan: ReturnType<typeof planPageSetup>): PageSetupSettings {
  if (!plan.ok) {
    throw new Error(`Kế hoạch không hợp lệ: ${plan.errors.join("; ")}`);
  }
  return plan.settings;
}

describe("planPageSetup: mặc định", () => {
  it("chuẩn Chính quyền theo file mẫu: A4 dọc, lề 20/20/30/20, header 12,5, footer 5,3 mm", () => {
    expect(settingsOf(planPageSetup(ND30))).toEqual({
      paperSize: "A4",
      orientation: "Portrait",
      marginsMm: { top: 20, bottom: 20, left: 30, right: 20 },
      headerDistanceMm: 12.5,
      footerDistanceMm: 5.3,
      mirrorMargins: false,
    });
  });

  it("chuẩn Đảng: lề 20/20/30/15, số trang cách mép trên 10 mm", () => {
    expect(settingsOf(planPageSetup(HD05))).toEqual({
      paperSize: "A4",
      orientation: "Portrait",
      marginsMm: { top: 20, bottom: 20, left: 30, right: 15 },
      headerDistanceMm: 10,
      footerDistanceMm: 12.5,
      mirrorMargins: false,
    });
  });

  it("giá trị mặc định của mỗi chuẩn đều hợp lệ với chính chuẩn đó", () => {
    for (const profile of [ND30, HD05]) {
      const { marginsMm } = settingsOf(planPageSetup(profile));
      for (const side of ["top", "bottom", "left", "right"] as const) {
        expect(isMarginValid(profile, side, marginsMm[side])).toBe(true);
      }
    }
  });
});

describe("planPageSetup: lề do người dùng nhập", () => {
  it("chấp nhận lề trong khoảng của chuẩn Chính quyền và giữ các lề không nhập ở mặc định", () => {
    const settings = settingsOf(planPageSetup(ND30, { marginsMm: { left: 35, right: 15 } }));
    expect(settings.marginsMm).toEqual({ top: 20, bottom: 20, left: 35, right: 15 });
  });

  it("từ chối lề ngoài khoảng, không áp dụng, nêu rõ cạnh, giá trị và khoảng cho phép", () => {
    const plan = planPageSetup(ND30, { marginsMm: { left: 25 } });
    expect(plan.ok).toBe(false);
    if (!plan.ok) {
      expect(plan.errors).toHaveLength(1);
      expect(plan.errors[0]).toContain("Lề trái 25 mm");
      expect(plan.errors[0]).toContain("30");
      expect(plan.errors[0]).toContain("35");
      expect(plan.errors[0]).toContain("Chính quyền (NĐ 30)");
    }
  });

  it("báo tất cả lề sai cùng lúc", () => {
    const plan = planPageSetup(ND30, { marginsMm: { top: 10, right: 40 } });
    expect(plan.ok).toBe(false);
    if (!plan.ok) {
      expect(plan.errors).toHaveLength(2);
      expect(plan.errors.join(" ")).toContain("Lề trên 10 mm");
      expect(plan.errors.join(" ")).toContain("Lề phải 40 mm");
    }
  });

  it("chuẩn Đảng cố định lề: lề phải 20 mm bị từ chối với thông báo \"đúng 15 mm\"", () => {
    const plan = planPageSetup(HD05, { marginsMm: { right: 20 } });
    expect(plan.ok).toBe(false);
    if (!plan.ok) {
      expect(plan.errors[0]).toContain("Lề phải 20 mm");
      expect(plan.errors[0]).toContain("đúng 15 mm");
    }
  });

  it("dùng dấu phẩy thập phân trong thông báo", () => {
    const plan = planPageSetup(ND30, { marginsMm: { top: 19.5 } });
    expect(plan.ok).toBe(false);
    if (!plan.ok) {
      expect(plan.errors[0]).toContain("Lề trên 19,5 mm");
    }
  });

  it("từ chối giá trị không phải số", () => {
    const plan = planPageSetup(ND30, { marginsMm: { top: Number.NaN } });
    expect(plan.ok).toBe(false);
    if (!plan.ok) {
      expect(plan.errors[0]).toContain("Lề trên");
      expect(plan.errors[0]).toContain("số hợp lệ");
    }
  });
});

describe("planPageSetup: in hai mặt", () => {
  it("chuẩn Đảng bật lề đối xứng, giữ lề trái 30 mm là lề trong và phải 15 mm là lề ngoài", () => {
    const settings = settingsOf(planPageSetup(HD05, { duplex: true }));
    expect(settings.mirrorMargins).toBe(true);
    expect(settings.marginsMm.left).toBe(30);
    expect(settings.marginsMm.right).toBe(15);
  });

  it("chuẩn Chính quyền không quy định in hai mặt nên từ chối", () => {
    const plan = planPageSetup(ND30, { duplex: true });
    expect(plan.ok).toBe(false);
    if (!plan.ok) {
      expect(plan.errors[0]).toContain("không quy định in hai mặt");
    }
  });

  it("không yêu cầu in hai mặt thì lề đối xứng tắt", () => {
    expect(settingsOf(planPageSetup(HD05, { duplex: false })).mirrorMargins).toBe(false);
  });
});

describe("parseMm", () => {
  it("đọc số có dấu chấm hoặc dấu phẩy thập phân, bỏ khoảng trắng", () => {
    expect(parseMm("20")).toBe(20);
    expect(parseMm(" 12,5 ")).toBe(12.5);
    expect(parseMm("12.5")).toBe(12.5);
  });

  it("trả về NaN với chuỗi rỗng hoặc không phải số", () => {
    expect(parseMm("")).toBeNaN();
    expect(parseMm("   ")).toBeNaN();
    expect(parseMm("abc")).toBeNaN();
    expect(parseMm("20mm")).toBeNaN();
    expect(parseMm("1,2,3")).toBeNaN();
  });
});

describe("findPageSetupMismatches", () => {
  const expected = settingsOf(planPageSetup(ND30));
  const actual = {
    widthPt: mmToPt(210),
    heightPt: mmToPt(297),
    topPt: mmToPt(20),
    bottomPt: mmToPt(20),
    leftPt: mmToPt(30),
    rightPt: mmToPt(20),
    headerDistancePt: mmToPt(12.5),
    footerDistancePt: mmToPt(5.3),
    mirrorMargins: false,
  };

  it("không có chênh lệch khi đọc lại đúng giá trị", () => {
    expect(findPageSetupMismatches(expected, actual)).toEqual([]);
  });

  it("chấp nhận sai số làm tròn của Word (dưới 0,1 mm; A4 lệch 0,03 mm)", () => {
    expect(
      findPageSetupMismatches(expected, {
        ...actual,
        widthPt: mmToPt(210.03),
        heightPt: mmToPt(297.02),
        leftPt: mmToPt(29.95),
      }),
    ).toEqual([]);
  });

  it("báo chênh lệch cụ thể bằng tiếng Việt khi lề hoặc khổ giấy sai", () => {
    const mismatches = findPageSetupMismatches(expected, {
      ...actual,
      leftPt: mmToPt(15),
      widthPt: mmToPt(216),
      mirrorMargins: true,
    });
    expect(mismatches).toHaveLength(3);
    expect(mismatches.join(" ")).toContain("lề trái");
    expect(mismatches.join(" ")).toContain("chiều rộng");
    expect(mismatches.join(" ")).toContain("lề đối xứng");
  });
});

describe("summarizeApply", () => {
  const settings = settingsOf(planPageSetup(HD05, { duplex: true }));

  it("báo thành công nêu khổ giấy, lề và số phần", () => {
    const result = summarizeApply(settings, 1, []);
    expect(result.kind).toBe("ok");
    expect(result.text).toBe("Đã thiết lập khổ A4, lề trên 20, dưới 20, trái 30, phải 15 mm, bật lề đối xứng, cho 1 phần.");
  });

  it("không nhắc lề đối xứng khi tắt, ghi số phần khi nhiều phần", () => {
    const plain = settingsOf(planPageSetup(ND30));
    expect(summarizeApply(plain, 3, []).text).toBe("Đã thiết lập khổ A4, lề trên 20, dưới 20, trái 30, phải 20 mm cho 3 phần.");
  });

  it("cảnh báo khi đọc lại có chênh lệch, liệt kê từng chênh lệch", () => {
    const result = summarizeApply(settings, 2, ["lề trái: đọc lại 15 mm, mong đợi 30 mm"]);
    expect(result.kind).toBe("warning");
    expect(result.text).toContain("không khớp");
    expect(result.text).toContain("lề trái: đọc lại 15 mm");
  });
});
